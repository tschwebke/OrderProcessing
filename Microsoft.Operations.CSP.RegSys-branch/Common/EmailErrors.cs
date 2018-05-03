using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        public static void EmailErrors(StringBuilder content, string subject, string errorName, string banner, WorkItem wi)
        {
            try
            {
                Dictionary<string, string> custom = new Dictionary<string, string>();

                custom.Add("###MAIN_CONTENT###", content.ToString());
                custom.Add("###BANNER###", banner);
                custom.Add("###NOTIFICATION_HYPERLINK###", string.Empty);
                custom.Add("###PRIVACY_LINK###", Resources.Privacy_Link);
                custom.Add("###TERMS_OF_USE###", Resources.TOU_Link);
                custom.Add("###TRADEMARKS###", "http://www.microsoft.com/about/legal/en/us/IntellectualProperty/Trademarks/Default.aspx");
                custom.Add("####.##.##.####", Assembly.GetExecutingAssembly().GetVersion());
                string bodyText = FileSystem.GetEmbeddedFileContent("Microsoft.Operations.CSP.RegSys.Templates.Invoice.html", "Microsoft.Operations.CSP.RegSys");
                string compressionResult = string.Empty;
                foreach (string s in custom.Keys)
                {
                    bodyText = bodyText.Replace(s, custom[s]);
                }

                ExchangeService svc = ExchangeServiceAutomatic.New(EmailUserIdentityEmail, EmailPassword);
                EmailMessage message = new EmailMessage(svc);

                bodyText = Optimize.MinifyHtml(bodyText, out compressionResult, true);
                message.Subject = string.Format(subject);

                if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
                {
                    message.ToRecipients.Add("chads@microsoft.com");
                    message.ToRecipients.Add("Svcasfp@microsoft.com");
                }
                else
                {
                    message.ToRecipients.Add("Svcasfp@microsoft.com");
                    //message.CcRecipients.Add("");
                }

                message.BccRecipients.Add("cboinfrastructure@microsoft.com");
                message.ReplyTo.Add("svcasfp@microsoft.com");
                message.Body = new MessageBody(BodyType.HTML, bodyText);
                Email.InsertImageFromResource(ref message, "microsoft_footer.png");
                Email.InsertImageFromResource(ref message, "asfp_banner.png");

                message.Save(WellKnownFolderName.Drafts); // this is required to get the "ID" value so we can access other properties of the object. After the mail gets sent, this is removed from DRAFTS

                message.Load(new PropertySet(ItemSchema.MimeContent));
                var mimeContent = message.MimeContent;

                string tempFileFolder = string.Format(@"{0}\{1}\Temp", FileSystem.BaseFolder, "CSP Workflow");
                if (!Directory.Exists(tempFileFolder)) Directory.CreateDirectory(tempFileFolder);

                var fileNamePayload = Path.Combine(tempFileFolder, string.Format("error_message_{0}.eml", errorName));

                using (var fileStream = new FileStream(fileNamePayload, FileMode.Create))
                {
                    fileStream.Write(mimeContent.Content, 0, mimeContent.Content.Length);
                }

                message.SendAndSaveCopy(WellKnownFolderName.SentItems);
                if (!string.IsNullOrEmpty(fileNamePayload))
                {
                    try
                    {
                        Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileNamePayload, "Error Email");
                        wi.Attachments.Add(eml);
                    }
                    catch (Exception ex)
                    {
                        wi.History += string.Format("[SYSTEM] Tried to attach email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                var emailFailedReason = string.Format("When attempting message delivery, a system error was encountered: '<i>{0}</i>'.", ex.Message);
                wi.History += string.Format("[SYSTEM] Tried to send an email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
            }
        }
    }
}