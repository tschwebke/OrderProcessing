using Microsoft.Exchange.WebServices.Data;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Reflection;
using System.Text;

namespace Microsoft.Operations
{
    /// <summary>
    /// Useful methods to assist with the construction of HTML email, using the
    /// Microsoft.Exchange.WebServices API.
    /// </summary>
    public static class Email
    {
        /// <summary>
        /// Forces the email being sent into 'debug mode', performs the following transforms:
        ///
        /// (a) Changes the specified recipient (b) Appends information and statistics about the
        /// outgoing email.
        /// </summary>
        /// <param name="email">A reference to the email being modified, prior to being sent.</param>
        /// <param name="isDebugMode">
        /// false = ignore this completely, true = include debug information and re-route
        /// </param>
        /// <param name="tester">
        /// Email address of the person who should received the email (instead) ... acts as an override.
        /// </param>
        /// <param name="additionalText">
        /// Anything additional that you need to add (e.g. from the process which constructed the original)
        /// </param>
        /// <remarks>
        /// This deliberately inserts the text outside the main html body, so that there is no need
        /// for maintaining markers in the text.
        /// </remarks>
        public static void AdjustDebugMode(ref EmailMessage email, bool isDebugMode, EmailAddress tester, string additionalText = default(string))
        {
            if (isDebugMode)
            {
                StringBuilder info = new StringBuilder();

                info.AppendLine("<br/><br/><table><tr><td><span style='color:gray;font-family:Courier New;'>");
                info.AppendLine("** DEBUG MODE **<br/><br/>");

                // info.AppendLine(string.Format("Microsoft BIOS Alerts {0}
                // <br/>
                // ", Assembly.GetCallingAssembly().GetVersion()));
                info.AppendLine(string.Format("Original Intended Recipient(s):<br/>"));

                foreach (EmailAddress recipient in email.ToRecipients)
                {
                    info.AppendLine(string.Format("To: {0} &lt;{1}&gt;<br/>", recipient.Name, recipient.Address));
                }

                if (!string.IsNullOrEmpty(additionalText))
                {
                    info.AppendLine(additionalText);
                }

                info.AppendLine(string.Format("sent using {0}<br/>", email.Service.Url));
                info.AppendLine("</span></td></tr></table>");

                // now override the recipient, send to nominated tester

                email.ToRecipients.Clear();
                email.ToRecipients.Add(tester);

                email.Body.Text = email.Body.Text.Insert(email.Body.Text.Length, info.ToString());
            }
        }

        /// <summary>
        /// Attempts to return an SMTP address, accepts any kind of dodgy input.
        /// 1) SMTP Address, just returns that if valid
        /// 2) DOMAIN\alias ... will return in the form of alias@microsoft.com
        /// 3) Display Name ... any other condition - performs an Active Directory lookup
        /// 4) NEW TFS 2015.1 format ... "Marcel Dorner &lt;EUROPE\\marceldo&gt;"
        /// </summary>
        public static string DetermineSMTPAddressFrom(string input)
        {
            string output = string.Empty;

            if (string.IsNullOrEmpty(input))
            {
                // do nothing
            }
            else if (input.IsValidEmailAddress())
            {
                output = input; // no conversion
            }
            else if (input.Contains("<") && input.Contains(@"\") && input.EndsWith(">"))
            {
                // This is a very specific format test for a TFS2015.1 value (yes, they changed
                // everything with this update). EXAMPLE is: "Marcel Dorner <EUROPE\\marceldo>"
                string[] identityInTwoParts = input.GetTextBetween("<", ">").Split(@"\".ToCharArray());
                output = ActiveDirectory.EmailAddressFromAliasDomain(identityInTwoParts[0].ToUpper(), identityInTwoParts[1]);
            }
            else if (input.Contains(@"\"))
            {
                string[] identityInTwoParts = input.Split(@"\".ToCharArray());
                output = ActiveDirectory.EmailAddressFromAliasDomain(identityInTwoParts[0].ToUpper(), identityInTwoParts[1]);
            }
            else
            {
                output = ActiveDirectory.EmailAddressFromDisplayName(input);
            }

            return output;
        }

        /// <author>Warren James (Adecco)</author>
        /// <summary>
        /// Replaces an image in the HTML, turns it into an embedded image with as few arguments and
        /// code as possible! The mechanic is an attachment, where the Exchange API handles the
        /// serialization (read below for exact detail) Intended to keep calling code neat and tidy,
        /// assumes the following conditions are true:
        ///
        /// (a) The HTML body content is already populated (i.e. call this last) (b) The calling
        /// assembly has a resource in the /Images folder (default). (c) Resource is marked as
        /// 'embedded content' (d) Resource access modifier is 'public' (internal is the default) -
        /// required so that the resources are visible
        ///
        /// This method is geared to reference the CALLING assembly to look for the embedded
        /// resource. The images themselves are embedded as a multipart resource in the internal
        /// MHTML generated. It does increase the size of the email however the benefit is that the
        /// images are instantly visible (i.e. no download required) ... and do not require external
        /// image/web hosting.
        /// </summary>
        /// <param name="email">
        /// Reference to the original message object, required since it modifies the body text inside.
        /// </param>
        /// <param name="imageFullStringBetweenQuotes">
        /// Name of the file, matching the item inside the html (e.g. "some_image_64.png")
        /// </param>
        /// <param name="alternateResourceLocation">
        /// If the image is not in the /images namespace/folder, you can specify its resource name.
        /// (fully qualified e.g. 'Microsoft.Operations.Webservices.Gallacake.Images.WinAzure_logo_Wht_rgb_D.png')
        /// </param>
        /// <remarks>
        /// A Resource name typically inherits the folder it resides in. Full syntax often looks
        /// something like the following pattern: "Microsoft.Calling.Assembly.Folder.Filename.ext" If
        /// you're unsure of the resources in your assembly, you can easily index them using a
        /// ResourceSet: http://stackoverflow.com/questions/2041000/loop-through-all-the-resources-in-a-resx-file
        /// </remarks>
        /// <todo>Need to add support for OTHER ASSEMBLIES</todo>
        public static void InsertImageFromResource(ref EmailMessage email, string imageFullStringBetweenQuotes, string alternateResourceLocation = default(string))
        {
            // The 'cid:#####' reference emulates the format of MHTML digest. The new name of the
            // image after the '@' character is arbitrary but should be unique to the content, hence
            // the random number usage.
            if (email.Body.Text.Contains(imageFullStringBetweenQuotes))
            {
                Assembly from = Assembly.GetCallingAssembly();
                string cidImage = string.Format("{0}@{1}", imageFullStringBetweenQuotes, Maths.RandomNumber(0, 1000000000));
                string embeddedResourceName = (string.IsNullOrEmpty(alternateResourceLocation)) ? string.Format("{0}.Images.{1}", from.FullName.Substring(0, from.FullName.IndexOf(',')), imageFullStringBetweenQuotes) : alternateResourceLocation;

                Attachment logo = email.Attachments.AddFileAttachment(imageFullStringBetweenQuotes, from.GetManifestResourceStream(embeddedResourceName).ReadFully());
                logo.ContentId = cidImage;
                logo.IsInline = true; // sets the mapi attribute: PR_ATTACHMENT_HIDDEN

                // Finally, adjust the MHTML so that the image reference now points to the internal
                // MIME attachement.

                email.Body.Text = email.Body.Text.Replace(imageFullStringBetweenQuotes, string.Format("cid:{0}", cidImage));
            }
        }

        /// <summary>
        /// Same image insertion, but using an Image already obtained.
        /// </summary>
        public static void InsertImageFromResource(ref EmailMessage email, System.Drawing.Image imageInRam, string imageFullStringBetweenQuotes)
        {
            if (email.Body.Text.Contains(imageFullStringBetweenQuotes))
            {
                string cidImage = string.Format("{0}@{1}", imageFullStringBetweenQuotes, Maths.RandomNumber(0, 1000000000));
                Microsoft.Exchange.WebServices.Data.Attachment logo = email.Attachments.AddFileAttachment(imageFullStringBetweenQuotes, imageInRam.ImageToByteArray(ImageFormat.Png));
                logo.ContentId = cidImage;
                logo.IsInline = true; // sets the mapi attribute: PR_ATTACHMENT_HIDDEN

                email.Body.Text = email.Body.Text.Replace(imageFullStringBetweenQuotes, string.Format("cid:{0}", cidImage));
            }
        }

        /// <summary>
        /// Quickly send an email to administor with a small/basic payload, from anywhere in the
        /// program. Runs in the security context of the currently running account.
        /// </summary>
        /// <param name="emailSender">
        /// Fully formed SMTP email address - denoting the mailbox in use, matching service account
        /// </param>
        /// <param name="subject">Basic subject line</param>
        /// <param name="body">Body content anything you want (is HTML-based).</param>
        public static void NotifyAdministrator(string emailSender, string subject, string body)
        {
            try
            {
                ExchangeService svc = ExchangeServiceAutomatic.New(emailSender);
                EmailMessage message = new EmailMessage(svc);
                message.Subject = subject;

                Assembly from = Assembly.GetCallingAssembly();

                string bodyHTML = FileSystem.GetEmbeddedFileContent("Microsoft.Operations.Templates.Blank.html", "Microsoft.Operations");
                bodyHTML = bodyHTML.Replace("###MAIN_CONTENT###", body);
                bodyHTML = bodyHTML.Replace("###NOTIFICATION_HYPERLINK###", "http://vstfpg07:8080/tfs/Operations/Online%20Operations/_workItems#id=19700&_a=edit");

                message.ToRecipients.Add("cboinfrastructure@microsoft.com");

                message.Body = new MessageBody(BodyType.HTML, bodyHTML);
                InsertImageFromResource(ref message, "microsoft_logo_300px.png");
                message.Send();
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Application", "Microsoft.Operations.v2016.common assembly, cannot notify administrator. Message was" + ex.Message + ex.InnerException + "and original message was '" + subject + ":" + body + "'");
            }
        }
    }
}