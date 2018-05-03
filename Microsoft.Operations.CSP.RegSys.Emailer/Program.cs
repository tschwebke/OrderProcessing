using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;

namespace Microsoft.Operations.CSP.RegSys.Emailer
{
    internal class Program
    {
        //public AzureLogger Log = new AzureLogger();

        public static ICredentials credentials = new NetworkCredential(EmailUserIdentityEmail.Replace("@microsoft.com", string.Empty), EmailPassword, "REDMOND");
        public static string EmailPassword = "Next200Wins!";
        public static string EmailUserIdentityEmail = "svcasfp@microsoft.com";
        public static string ProjectName = "CSP";
        public static string ServerName = "http://vstfpg07:8080/tfs/Operations";

        public static void EmailGeneric(WorkItem wi, RegSysWizardItem data)
        {
            string fileNamePayload = string.Empty;
            StringBuilder textSummary = new StringBuilder();
            Dictionary<string, string> custom = new Dictionary<string, string>();
            data.SKUs = data.SKUs.Replace("\r\n", ";");
            string[] skusArray = data.SKUs.Split(';');

            bool bPromo = data.SpecialInstructions.Contains("PATN");

            custom.Add("###BANNER###", "Microsoft Advanced Support for Partners Confirmation");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #'>Dear {0},</span><br/>", data.PartnerContactName.Trim());
            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            if (bPromo)
            {
                textSummary.AppendLine("Thank you for submitting an order request for Microsoft Advanced Support for Partners (ASfP) Transition Package!<br/></span>");
            }
            else
            {
                textSummary.AppendLine("Thank you for submitting an order request for Microsoft Advanced Support for Partners (ASfP)!<br/></span>");
            }

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine("You will find your order details below:</span>");

            textSummary.AppendLine("<ul type='disc'>");

            if (bPromo)
            {
                textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
                textSummary.AppendLine(@"Advanced Support for Partners Transition Package – Promocode: {0}", data.SpecialInstructions);
                textSummary.AppendLine("</span></li>");
            }

            if (bPromo)
            {
                foreach (string sku in skusArray)
                {
                    if (sku != "")
                    {
                        bool bNoSKU = sku.Contains("Advanced Support for Partners Monthly Subscription");
                        bool bMonthly = sku.Contains("No SKU items appear");
                        if (!bNoSKU && !bMonthly)
                        {
                            textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
                            textSummary.AppendLine(@"{0}", sku);
                            textSummary.AppendLine("</span></li>");
                        }
                    }
                }
            }
            else
            {
                foreach (string sku in skusArray)
                {
                    if (sku != "")
                    {
                        textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
                        textSummary.AppendLine(@"{0}", sku);
                        textSummary.AppendLine("</span></li>");
                    }
                }
            }

            textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
            if (data.PartnerArea != "")
            {
                textSummary.AppendLine(@"Preferred Services Account Manager Location: {0}", data.PartnerArea);
            }
            else
            {
                textSummary.AppendLine(@"Preferred Services Account Manager Location: (no prefered location specified)");
            }
            textSummary.AppendLine("</span></li>");

            if (!bPromo)
            {
                textSummary.AppendLine("</span></li>");
                textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
                if (data.SpecialInstructions != "")
                {
                    textSummary.AppendLine(@"Special Instructions: {0}", data.SpecialInstructions);
                }
                else
                {
                    textSummary.AppendLine(@"Special Instructions: (no special instructions provided)");
                }
            }

            textSummary.AppendLine("</span></li>");

            textSummary.AppendLine("</ul>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");

            if (bPromo)
            {
                textSummary.AppendLine("A Microsoft representative will be contacting you within 72 hours (Business Hours: Monday to Friday 9am– 5pm Pacific Time Zone) to finalize and activate your order, including any Microsoft Dynamics-specific benefits that come with your transition package.<br/></span>");
            }
            else
            {
                textSummary.AppendLine("A Microsoft representative will be contacting you within 72 hours (Business Hours: Monday to Friday 9am– 5pm Pacific Time Zone) to finalize the order and activate ASfP for your company!<br/></span>");
            }

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine("In the meantime, please feel free to reach out to Advanced Support for Partners at <a href='mailto:svcasfp@microsoft.com'>svcasfp@microsoft.com</a> with any questions.<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine("Thank you for your continued partnership and support in delivering Microsoft technologies to our mutual customers.<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine("Sincerely,<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>");
            textSummary.AppendLine("Your Microsoft Advanced Support for Partners team<br/></span>");

            //custom.Add("style='background-color:#111111;", "style='background-color:#3399FF;");
            // custom.Add("width:320px;height:106px", "width:320px;height:106px");

            // These substitutions are common

            custom.Add("###NOTIFICATION_HYPERLINK###", string.Empty);
            custom.Add("###MAIN_CONTENT###", textSummary.ToString());
            custom.Add("###PRIVACY_LINK###", "http://aka.ms/asfp_privacy");
            custom.Add("###TERMS_OF_USE###", "http://aka.ms/asfp_agreement");
            custom.Add("###TRADEMARKS###", "http://www.microsoft.com/about/legal/en/us/IntellectualProperty/Trademarks/Default.aspx");
            custom.Add("####.##.##.####", Assembly.GetExecutingAssembly().GetVersion());
            custom.Add("banner_graphic.png", "asfp_banner.png");

            ///////////////////////////////////////////////////////////////////////////////////////////
            // System Notification #####

            string bodyText = FileSystem.GetEmbeddedFileContent("Microsoft.Operations.CSP.RegSys.Templates.ASfPColorful.html", "Microsoft.Operations.CSP.RegSys");
            string compressionResult = string.Empty;

            foreach (string s in custom.Keys)
            {
                bodyText = bodyText.Replace(s, custom[s]);
            }

            bodyText = Optimize.MinifyHtml(bodyText, out compressionResult, true);

            ExchangeService svc = ExchangeServiceAutomatic.New(EmailUserIdentityEmail, EmailPassword);
            EmailMessage message = new EmailMessage(svc);

            message.Subject = string.Format("Microsoft Advanced Support for Partners (ASfP) Offering Purchase Confirmation");
            // message.ToRecipients.Add("a-wjames@microsoft.com"); // for testing only

            message.ToRecipients.Add(data.PartnerContactEmail);

            if (!string.IsNullOrEmpty(data.BillingContactEmail))
            {
                string[] addr = data.BillingContactEmail.Split(';');
                foreach (string s in addr)
                {
                    if (s.IsValidEmailAddress())
                    {
                        message.CcRecipients.Add(s);
                    }
                }
            }

            message.BccRecipients.Add("cboinfrastructure@microsoft.com");
            message.BccRecipients.Add("asfpsales@microsoft.com");
            message.ReplyTo.Add("svcasfp@microsoft.com");
            message.Body = new MessageBody(BodyType.HTML, bodyText);

            Email.InsertImageFromResource(ref message, "microsoft_footer.png");
            Email.InsertImageFromResource(ref message, "asfp_banner.png");

            message.Save(WellKnownFolderName.Drafts); // this is required to get the "ID" value so we can access other properties of the object. After the mail gets sent, this is removed from DRAFTS

            message.Load(new PropertySet(ItemSchema.MimeContent));
            var mimeContent = message.MimeContent;

            string tempFileFolder = string.Format(@"{0}\{1}\Temp", FileSystem.BaseFolder, "CSP Workflow");
            if (!Directory.Exists(tempFileFolder)) Directory.CreateDirectory(tempFileFolder);

            fileNamePayload = Path.Combine(tempFileFolder, string.Format("purchase_confirmation_{0}.eml", data.PartnerContactName.RemoveWhitespace().RemoveInvalidFileNameCharacters().ToLower().MaxLength(5)));

            using (var fileStream = new FileStream(fileNamePayload, FileMode.Create))
            {
                fileStream.Write(mimeContent.Content, 0, mimeContent.Content.Length);
            }

            // Attach the file to the work item and write the history ...

            // message.Send();
            message.SendAndSaveCopy(WellKnownFolderName.SentItems);

            if (!string.IsNullOrEmpty(fileNamePayload))
            {
                // Also add, as an attachment, original spreadsheet submission.
                Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileNamePayload, "Confirmation Email");
                wi.Attachments.Add(eml);
            }

            if (wi.ValidatesOkay())
            {
                wi.History += "[SYSTEM] This record was identified as not having had a confirmation email sent. This is a reissue using the same content originally received. For more information about this capability (and context), please contact cboinfrastructure@microsoft.com.";
                wi.Save();
            }

            Console.WriteLine("FINISHED!");
            Console.WriteLine("Press any key to continue");
            PromptForItemNumber(); // repeat again
        }

        public static void Finished(bool isAborted = false)
        {
            if (isAborted)
            {
                Console.WriteLine("... exiting at request of user.");
            }
            else
            {
                Console.WriteLine("Press any key to exit");
            }
            Console.ReadKey();
        }

        // next change will be March 2017, but hopefully we can eliminate from source code by that time.
        private static void Main(string[] args)
        {
            Console.WriteLine("This console application is used for resend of the SKU confirmation email.");
            Console.WriteLine("USE WITH CARE (i.e. delivers customer-facing payload)");
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
            PromptForItemNumber();
        }

        /// <summary>
        /// Part 2, actually do the lookup / load details and get user to confirm that this is indeed
        /// the one.
        /// </summary>
        private static void ProcessTfsItem(int id)
        {
            TfsTeamProjectCollection tfs = null;
            WorkItemStore workItemStore = null;

            tfs = new TfsTeamProjectCollection(new Uri(ServerName), credentials);
            tfs.EnsureAuthenticated();
            workItemStore = new WorkItemStore(tfs);
            WorkItem wi = workItemStore.GetWorkItem(id);

            if (wi == null)
            {
                Console.WriteLine();
                Console.WriteLine("That work item doesn't appear to exist! (please check)");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
                PromptForItemNumber();
            }
            else
            {
                // normal processing!
                Console.WriteLine();
                RegSysWizardItem item = new RegSysWizardItem();

                item.PartnerContactName = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Name");
                item.PartnerContactEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Email");
                item.SKUs = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs");
                item.PartnerArea = wi.GetFieldValue("Microsoft.Operations.Partners.AreaOfExpertise");
                item.SpecialInstructions = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.Notes");
                item.BillingContactEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactEmail");

                Console.WriteLine("Name: " + item.PartnerContactName);
                Console.WriteLine("Email: " + item.PartnerContactEmail);
                Console.WriteLine("-------------------------------------------------------------");
                Console.WriteLine("SKUS: " + item.SKUs);
                Console.WriteLine();
                Console.WriteLine("Is this correct? (y/n)");
                string answer = Console.ReadLine();
                if (answer == "y")
                {
                    EmailGeneric(wi, item);
                    Finished();
                }
                else
                {
                    Finished(true);
                }
            }
        }

        /// <summary>
        /// Part 1, obtain a valid number to lookup.
        /// </summary>
        private static void PromptForItemNumber()
        {
            Console.Clear();
            Console.WriteLine("Please enter the TFS ID of the 'SKU Purchase' item you wish to re-issue:");
            string input = Console.ReadLine();

            int tfs_identifier;
            if (int.TryParse(input, out tfs_identifier))
            {
                Console.WriteLine("Loading SKU Purchase " + tfs_identifier + "...");
                ProcessTfsItem(tfs_identifier);
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Not a valid number!");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
                PromptForItemNumber();
            }
        }

        // end EmailSpreadsheetProcessingResult
    }
}