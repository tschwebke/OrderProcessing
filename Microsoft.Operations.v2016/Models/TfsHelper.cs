using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml;

namespace Microsoft.Operations
{
    public static class TfsHelper
    {
        /// <summary>
        /// TODO: These should be in common TFS library
        /// </summary>
        public static string GetCoreFieldValue(XmlElement eventData, string section, string target, string referenceName)
        {
            return GetCoreFieldValue(eventData, section, target, referenceName, "StringFields");
        }

        public static string GetCoreFieldValue(XmlElement eventData, string section, string target, string referenceName, string fieldDataType)
        {
            string referenceNameMatch = string.Format("{0}/{3}/Field[ReferenceName='{1}']/{2}", section, referenceName, target, fieldDataType);
            XmlNode nodeToFind;

            // NOTE: An example of the Pattern we want to find is eventData.SelectSingleNode("CoreFields/StringFields/Field[System.Title]/OldValue);

            nodeToFind = eventData.SelectSingleNode(referenceNameMatch);

            if (nodeToFind != null)
            {
                // It was found, manipulate it.
                return nodeToFind.InnerText;
            }
            else
            {
                // It was not found.
                return string.Empty;
            }
        }

        /// <summary>
        /// Examines the (string) contents of a specified (string) field and returns true if it
        /// appears the item has changed. i.e. between NewValue and OldValue. NOTE: If the item can't
        /// be found then it will return false.
        /// </summary>
        /// <param name="section">"CoreFields" or "ChangedFields" (only)</param>
        public static bool IsFieldModified(XmlElement eventData, string section, string referenceName)
        {
            // TODO: Eventually this might be condensed into a common routine, or there might also be
            // a TeamFoundationServer routine which is a bit more efficient, which could be used instead.

            bool isChanged = false;

            if (!string.IsNullOrEmpty(referenceName))
            {
                string fieldLocationOld = string.Format("{0}/StringFields/Field[ReferenceName='{1}']/OldValue", section, referenceName);
                string fieldLocationNew = string.Format("{0}/StringFields/Field[ReferenceName='{1}']/NewValue", section, referenceName);

                XmlNode nodeToFindOld = eventData.SelectSingleNode(fieldLocationOld);
                XmlNode nodeToFindNew = eventData.SelectSingleNode(fieldLocationNew);

                if (nodeToFindOld != null && nodeToFindNew != null)
                {
                    if (nodeToFindOld.InnerText != nodeToFindNew.InnerText)
                    {
                        return true; // Nodes were found, result comes from direct comparison of the two values.
                    }
                    else
                    {
                        return false; // the values were there, but they appear to be the same.
                    }
                }
                else
                {
                    return false; // The nodes could not be found, therefore no 'changed' comparison.
                }
            }

            return isChanged;
        }

        /// <summary>
        /// NOTE: This is an expensive operation, so where possible
        /// </summary>
        /// <param name="tfsUri">
        /// url of the server, e.g. "http://vstfpg07:8080/tfs/Operations", this is not project sensitive.
        /// </param>
        /// <param name="groupName">Name of TFS Group or Team you want to iterate members of.</param>
        public static Dictionary<string, string> LookupGroupMembership(string tfsUri, string groupName)
        {
            Dictionary<string, string> members_tfs = new Dictionary<string, string>();
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsUri));

            tfs.EnsureAuthenticated();

            IIdentityManagementService iims = tfs.GetService<IIdentityManagementService>();
            TeamFoundationIdentity sids = iims.ReadIdentity(IdentitySearchFactor.DisplayName, groupName, MembershipQuery.Expanded, ReadIdentityOptions.None);
            TeamFoundationIdentity[] target_folk = iims.ReadIdentities(sids.Members, MembershipQuery.Expanded, ReadIdentityOptions.None);

            // EventLog.WriteEntry("Application", string.Format("[LookupGroupMembership] Total of {0}
            // detected in '{1}'", target_folk.Count(), groupName));

            foreach (TeamFoundationIdentity member in target_folk)
            {
                if (!member.IsContainer && member.DisplayName != "BIOS_OPS_Applications" && member.DisplayName != "Microsoft CSP Submissions")
                {
                    members_tfs.AddOrUpdate(member.DisplayName, member.GetAttribute("Mail", string.Empty).ToLower()); // add directly
                }
            }

            return members_tfs;
        }

        /// <summary>
        /// Identifies an email address based on a user in Team Foundation Server
        /// </summary>
        /// <param name="tfsUri">Collection information / URL</param>
        /// <param name="userDisplayName">Full display name of the person you wish to find</param>
        /// <returns></returns>
        public static Dictionary<string, string> LookupUserEmail(string tfsUri, string userDisplayName)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsUri));
            tfs.EnsureAuthenticated();
            IIdentityManagementService iims = tfs.GetService<IIdentityManagementService>();
            TeamFoundationIdentity person = iims.ReadIdentity(IdentitySearchFactor.DisplayName, userDisplayName, MembershipQuery.None, ReadIdentityOptions.None);

            Dictionary<string, string> person_contact = new Dictionary<string, string>();
            person_contact.Add(userDisplayName, person.GetAttribute("Mail", string.Empty));
            return person_contact;
        }

        /// <summary>
        /// TODO: Requires unification
        /// </summary>
        public static WorkItem RetrieveEntireWorkItem(WorkItemStore wis, int workItemID)
        {
            // Extranet instances only allow a https (secure) connection, but sometimes this is not
            // specified by the internals of the TFS event mechanism. So, we make an allowance for it here.

            try
            {
                WorkItem wi = wis.GetWorkItem(workItemID);
                return wi;
            }
            catch (Exception ex)
            {
                // can't find it, or the item doesn't exist ... or whatever.
                EventLog.WriteEntry("Application", ex.Message);
                return null;
            }
        }
    }
}