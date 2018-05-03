using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using System.ServiceModel;
using System.Xml.Linq;

namespace Microsoft.Operations
{
    /// <summary>
    /// Shorthand syntax for referencing Active Directory property bag values
    /// </summary>
    public enum AccountProperty
    {
        /// <summary>
        /// Fully-formed SMTP address
        /// </summary>
        EmailAddress,

        /// <summary>
        /// Common name of object, as often seen in Outlook, usually applicable to all microsoft.com objects
        /// </summary>
        DisplayName,

        /// <summary>
        /// 8 characters or less, typically lowercase
        /// </summary>
        AccountAlias
    }

    /// <summary>
    /// Contains code which interacts with Active Directory - so you don't have to!!
    /// </summary>
    public static class ActiveDirectory
    {
        /// <summary>
        /// Returns the display name of the current user, as identified by the WCF service in
        /// relation to the caller. i.e. the WCF automatically determines the user alias and uses it
        /// for a lookup in Active Directory.
        /// </summary>
        public static string DisplayName()
        {
            string[] userIdentity = OperationContext.Current.ServiceSecurityContext.WindowsIdentity.Name.Split("\\".ToCharArray());
            string displayName = string.Empty;

            if (userIdentity.Length > 0)
            {
                string domainName = userIdentity[0].ToString();
                string userAlias = userIdentity[1].ToString();
                displayName = GlobalDirectoryLookupUserAccountProperty("displayName", "SAMAccountName", userAlias, string.Empty);
            }
            return displayName;
        }

        /// <summary>
        /// Performs a lookup in active directory, for a record, given the alias. Returns the display
        /// name (friendly name). Optionally supply a DOMAIN in the forest for faster performance
        /// (preferred if the domain is known). If a user wasn't found, test whether or not the name
        /// belongs to a distribution group instead.
        /// </summary>
        /// <param name="userAlias">The user alias of the user to lookup.</param>
        /// <param name="domainName">The domain which we expect (know) the user to be on.</param>
        public static string DisplayNameFromAlias(string userAlias, string domainName)
        {
            string displayName = GlobalDirectoryLookupUserAccountProperty("displayName", "SAMAccountName", userAlias, domainName);
            if (string.IsNullOrEmpty(displayName)) { displayName = LookupSecurityGroup(string.Empty, userAlias, AccountProperty.DisplayName); }
            return displayName;
        }

        /// <summary>
        /// Allows supplying the identity information in the form of DOMAIN\alias, which is a
        /// shorthand method. requires the presence of the slash, otherwise just returns an empty string.
        /// </summary>
        /// <param name="aliasIncludingDomainNamePrefix">
        /// Identity of the user, in DOMAIN\alias format.
        /// </param>
        public static string EmailAddressFromAliasDomain(string aliasIncludingDomainNamePrefix)
        {
            if (aliasIncludingDomainNamePrefix.Contains(@"\"))
            {
                string[] identityInTwoParts = aliasIncludingDomainNamePrefix.Split(@"\".ToCharArray());
                return EmailAddressFromAliasDomain(identityInTwoParts[0].ToUpper(), identityInTwoParts[1]);
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the 'mail' attribute from the Active Directory lookup, given an account name.
        /// </summary>
        public static string EmailAddressFromAliasDomain(string domainName, string displayName)
        {
            return GlobalDirectoryLookupUserAccountProperty("mail", "SAMAccountName", displayName, domainName);
        }

        /// <summary>
        /// Returns the 'email' address (alias) of a active directory object, given only the display
        /// name. Acheives this by performing a lookup on Active Directory, Use this method where the
        /// email address is unknown. This function will also automatically check security group
        /// names (but only if it can't find a user match).
        /// EDIT: use 'mail' if you want the full email address
        /// </summary>
        /// <param name="displayName">The display name of the object, as found in Active Directory</param>
        /// <param name="domainName">
        /// Supply this value if it is known, otherwise we'll try find them based on
        /// </param>
        public static string EmailAddressFromDisplayName(string displayName, string domainName = "")
        {
            // string emailAddress = ActiveDirectory.GlobalDirectoryLookupUserAccountProperty("mail",
            // "displayname", displayName, domainName);

            string emailAddress = string.Empty;
            string aliasUser = GlobalDirectoryLookupUserAccountProperty("SAMAccountName", "displayname", displayName, domainName);

            // typically a Microsoft corp address, if returned a match!
            if (!string.IsNullOrEmpty(aliasUser))
            {
                emailAddress = string.Format("{0}@microsoft.com", aliasUser);
            }
            else if (string.IsNullOrEmpty(aliasUser) && !displayName.Contains(" "))
            {
                // this may indeed be,
                emailAddress = EmailAddressFromAliasDomain(domainName, displayName);
            }

            // finally ... try security group ...

            if (string.IsNullOrEmpty(emailAddress)) { emailAddress = LookupSecurityGroup(displayName, string.Empty, AccountProperty.EmailAddress); }
            return emailAddress;
        }

        /// <summary>
        /// Gets the user thumbnail photo from active directory account.
        /// </summary>
        /// <param name="SAMAccountName">Security Accounts Manager account name (alias).</param>
        /// <returns>Bitmap pictures in bytes.</returns>
        public static byte[] GetUserThumbnailPhotoFromAccount(string SAMAccountName)
        {
            byte[] thumbnailInBytes = null;

            using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain))
            {
                using (UserPrincipal userPrincipal = new UserPrincipal(principalContext))
                {
                    userPrincipal.SamAccountName = SAMAccountName;
                    using (PrincipalSearcher principalSearcher = new PrincipalSearcher())
                    {
                        principalSearcher.QueryFilter = userPrincipal;
                        Principal principal = principalSearcher.FindOne();
                        if (principal != null)
                        {
                            DirectoryEntry directoryEntry = (DirectoryEntry)principal.GetUnderlyingObject();
                            PropertyValueCollection collection = directoryEntry.Properties["thumbnailPhoto"];

                            if (collection.Value != null && collection.Value is byte[])
                                thumbnailInBytes = (byte[])collection.Value;
                        }
                    }
                }
            }

            return thumbnailInBytes;
        }

        /// <summary>
        /// Universal lookup method which searches for a person/user object in active directory
        /// across ALL forests. Errors will silently fail with a zero-length string returned.
        /// IMPORTANT NOTES:
        /// 1. This search procedure returns the first found result (i.e. no support for multiples).
        /// 2. The named properties named must be the same format as used in LDAP searches.
        /// 3. The executing account for the web service (aka the Identity of the application pool)
        ///    needs to be a DOMAIN account - i.e. if it is running as 'LocalSystem' or
        ///    'NetworkService' then it won't work.
        /// </summary>
        /// <param name="outputProperty">
        /// The property to load from the results (does not support multiples).
        /// </param>
        /// <param name="searchProperty">
        /// The property we will be searching on (likewise, single value only) e.g. "SAMAccountName"
        /// for Alias.
        /// </param>
        /// <param name="searchValue">The actual input - e.g. someone's user alias.</param>
        /// <param name="domainName">
        /// OPTIONAL: If known, this can help improve the search time by only searching a specific
        ///           domain. Use string.Empty if not known.
        /// </param>
        public static string GlobalDirectoryLookupUserAccountProperty(string outputProperty, string searchProperty, string searchValue, string domainName)
        {
            DirectorySearcher searcher;
            string forest = string.Empty;

            // Because the domain may not be supplied, we have to check BOTH locations. Technically
            // this could be done as a single entry, but I do not know how to do it (without
            // searching the entire directory which proves to be a large performance hit) If we can't
            // find the user in CORP THEN we will automatically look in EXTRANET, following.

            forest = (domainName == "PARTNERS") ? "extranet" : "corp";

            if (!string.IsNullOrEmpty(domainName))
            {
                searcher = new DirectorySearcher(new DirectoryEntry(string.Format("GC://{0}.{1}.microsoft.com/DC={0},DC={1},DC=microsoft,DC=com", domainName, forest)));
                searcher.Filter = string.Format("(&({0}={1})(objectCategory=person)((objectClass=user)))", searchProperty, searchValue);
                searcher.PropertiesToLoad.Add(outputProperty);
                searcher.SearchScope = SearchScope.Subtree;
            }
            else
            {
                // if not specified then just use corp by default.
                searcher = new DirectorySearcher(new DirectoryEntry("GC://corp.microsoft.com/DC=corp,DC=microsoft,DC=com"));
                searcher.Filter = string.Format("(&({0}={1})(objectCategory=person)((objectClass=user)))", searchProperty, searchValue);
                searcher.PropertiesToLoad.Add(outputProperty);
                searcher.SearchScope = SearchScope.Subtree;
            }

            try
            {
                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    return result.Properties[outputProperty][0].ToString();
                }
                else
                {
                    // no results there, so try extranet !
                    // SearchResult result2 = searcher_extranet.FindOne();
                    //SearchResultCollection result2 = searcher_extranet.FindAll();

                    //if (result2 != null)
                    //{
                    //    foreach (SearchResult sr in result2)
                    //    {
                    //        String test = sr.Properties["adspath"][0].ToString();
                    //        if(test.Contains(searchValue))
                    //        {
                    //            return sr.Properties[outputProperty][0].ToString();
                    //        }
                    //    }
                    //    return string.Empty;
                    //}
                    //else
                    //{
                    //    // give up
                    //    return string.Empty;
                    //}

                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Application", string.Format("{0}|{1}", ex.Message, ex.StackTrace));
                return string.Empty;
            }
            finally
            {
                searcher.Dispose();
            }
        }

        /// <summary>
        /// Returns information about group membership for a given user. Works with multiple groups
        /// instead of single.
        /// </summary>
        /// <param name="delimitedGroupNames">
        /// PIPE delimited values, group names. Ensure that each is in DOMAIN\name format.
        /// </param>
        public static XElement IsCurrentUserMemberOfAnyOfTheFollowingGroups(string delimitedGroupNames)
        {
            XElement errorCollection = new XElement("Errors");
            XElement groupListing = new XElement("Groups");
            int highestAuthorizationLevel = 0;

            WindowsIdentity userIdentity = OperationContext.Current.ServiceSecurityContext.WindowsIdentity;

            if (userIdentity == null)
            {
                errorCollection.Add(new XElement("Error", string.Format("The Security context of the connecting user could not be identified: {0}", userIdentity.Name)));
            }
            else
            {
                WindowsPrincipal userPrincipal = new WindowsPrincipal(userIdentity);
                string[] groups = delimitedGroupNames.Split('|');

                foreach (string g in groups)
                {
                    int result = IsUserMemberOfGroup(g, userPrincipal);
                    XElement groupResult = new XElement("Group", g);
                    groupResult.SetAttributeValue("auth", result);
                    groupListing.Add(groupResult);

                    if (result > highestAuthorizationLevel) { highestAuthorizationLevel = result; }
                }
            }

            XElement xmlResponse = new XElement("IsCurrentUserMemberOfAnyOfTheFollowingGroupsResponse",
                new XElement("AuthorizationLevel", highestAuthorizationLevel),
                new XElement("CurrentUser", userIdentity.Name),
                groupListing,
                errorCollection
            );

            return xmlResponse;
        }

        /// <summary>
        /// Simple access to do a generic check for whether a user is a member of a security group.
        /// This information is freely available from the corpnet active directory (it's not secret),
        /// however this particular function checks the windows identity of the CURRENT CONNECTING
        /// USER only. Various error return codes:
        /// -1 = Not a member of group 0 = Unknown (failure to check / could not check) 1 = Is a
        ///  member of specified group.
        /// </summary>
        /// <param name="groupName">
        /// The groupname must be in DOMAIN\alias format. If the group's display name is given, it
        /// must first be converted.
        /// </param>
        public static int IsCurrentUserMemberOfGroup(string groupName)
        {
            // If the name of the group is supplied as an alias only, then we need to find the name
            // of the group

            if (!groupName.Contains("\\"))
            {
                groupName = string.Format("REDMOND\\{0}", LookupSecurityGroup(groupName, string.Empty, AccountProperty.AccountAlias));
            }

            WindowsIdentity userIdentity = OperationContext.Current.ServiceSecurityContext.WindowsIdentity;
            if (userIdentity == null) { return 0; }
            else
            {
                WindowsPrincipal userPrincipal = new WindowsPrincipal(userIdentity);
                return IsUserMemberOfGroup(groupName, userPrincipal);
            }
        }

        /// <summary>
        /// Quick check to determine whether a user is a member of a particular security group or
        /// not. This includes Active Directory chained membership - e.g. is a member of group x,
        /// which in turn is a member of y, etc. IMPORTANT NOTE: There are cases where IsInRole does
        /// not work sometimes
        /// TODO: Implement what is written in this link: http://stackoverflow.com/questions/323831/windowsprincipal-isinrole-and-universal-vs-global-active-directory-groups
        /// </summary>
        /// <param name="fullSecurityGroupName">The name of the security group.</param>
        /// <param name="windowsUser">
        /// Active Directory oject that you wish to perform the lookup on
        /// </param>
        public static int IsUserMemberOfGroup(string fullSecurityGroupName, WindowsPrincipal windowsUser)
        {
            int result = 0;

            if (windowsUser.IsInRole(fullSecurityGroupName)) { result = 1; }
            else { result = -1; }

            return result;
        }

        /// <summary>
        /// Security group lookup requires a different type of object for searching. This function
        /// will return multiple results if applicable according to the search pattern (do with what
        /// you want).
        /// NOTE: Security groups are created on the REDMOND domain (so the domain specification is hardcoded).
        /// Typically, use either one of the input parameters, with a string.Empty for the other value.
        /// </summary>
        /// <param name="displayName">
        /// display name for lookup - use this when the group is known but not the alias.
        /// </param>
        /// <param name="emailAlias">
        /// The alias by which the security group is known, not the full email address.
        /// </param>
        /// <param name="outputProperty">
        /// Indicate which property is desired (choice based on enumerator, saves having to know the syntax)
        /// </param>
        public static string LookupSecurityGroup(string displayName, string emailAlias, AccountProperty outputProperty)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "redmond");
            GroupPrincipal g = new GroupPrincipal(ctx);
            String foundValue = string.Empty;

            if (!string.IsNullOrEmpty(displayName)) { g.DisplayName = displayName; } // e.g. "APOC BIOS VSTF Support";
            if (!string.IsNullOrEmpty(emailAlias)) { g.SamAccountName = emailAlias; } // e.g. "biosvstf"

            PrincipalSearcher ps = new PrincipalSearcher(g);
            PrincipalSearchResult<Principal> results = ps.FindAll();

            if (results != null)
            {
                foreach (Principal p in results)
                {
                    switch (outputProperty)
                    {
                        case AccountProperty.AccountAlias: foundValue = p.SamAccountName; break;
                        case AccountProperty.DisplayName: foundValue = p.DisplayName; break;
                        case AccountProperty.EmailAddress: foundValue = p.SamAccountName + "@microsoft.com"; break;
                    }
                    break; // only process the first entry
                }
            }

            results.Dispose();
            g.Dispose();
            ctx.Dispose();

            return foundValue;
        }

        /// <summary>
        /// Takes a given (redmond-domain-based) security group name, and returns the users (raw) who
        /// are found to be members of that group.
        /// NOTE: Only mail-enabled members will be shown, since it's assumed that we wish to use
        ///       this for email distribution.
        /// </summary>
        public static Dictionary<string, string> LookupSecurityGroupMembers(string displayName, AccountProperty outputProperty)
        {
            Dictionary<string, string> member_output = new Dictionary<string, string>();
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "redmond");
            GroupPrincipal g = new GroupPrincipal(ctx);

            try
            {
                if (!string.IsNullOrEmpty(displayName))
                {
                    g.DisplayName = displayName;
                }

                PrincipalSearcher ps = new PrincipalSearcher(g);
                GroupPrincipal result = (GroupPrincipal)ps.FindOne();

                if (result != null)
                {
                    foreach (Principal p in result.Members)
                    {
                        member_output.AddOrUpdate(p.DisplayName, string.Format("{0}@microsoft.com", p.SamAccountName));
                    }
                }

                g.Dispose();
                ctx.Dispose();
            }
            catch (Exception ex)
            {
                // sometimes Active Directory screws up !
                EventLog.WriteEntry("Application", string.Format("{0}|{1}", ex.Message, ex.StackTrace));
            }

            return member_output;
        }
    }
}