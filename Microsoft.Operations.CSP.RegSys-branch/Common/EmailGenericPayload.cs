﻿using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        public static void EmailGeneric(RegSysWizardItem data, out string fileNamePayload, string mailLang)
        {
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(mailLang);
            StringBuilder textSummary = new StringBuilder();
            Dictionary<string, string> custom = new Dictionary<string, string>();
            data.SKUs = data.SKUs.Replace("apply),5-Pack", "apply)\r\n5-Pack"); // This was to handle some earlier bad data
            data.SKUs = data.SKUs.Replace("31,800 円 (税別),5 インシデント パック", "31,800 円 (税別)\r\n5 インシデント パック");
            data.SKUs = data.SKUs.Replace("\r\n", ";");
            string[] skusArray = data.SKUs.Split(';');

            bool bPromo = data.SpecialInstructions.Contains("PATN");

            custom.Add("###BANNER###", Resources.Banner);
            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #'>" + Resources.Dear + "</span><br/>", data.PartnerContactName.Trim());
            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");

            if (bPromo)
            {
                textSummary.AppendLine(Resources.Submit_ThankYou_PA + "<br/></span>");
            }
            else
            {
                textSummary.AppendLine(Resources.Submit_ThankYou_ASfP + "<br/></span>");
            }

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine(Resources.Submit_OrderDetails + "</span>");
            textSummary.AppendLine("<ul type='disc'>");

            if (bPromo)
            {
                textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>" + Resources.Submit_PA_PromoCode + "</span></li>", data.SpecialInstructions);
            }

            if (bPromo)
            {
                foreach (string sku in skusArray)
                {
                    if (sku != "")
                    {
                        bool bNoSKU = sku.Contains("Advanced Support for Partners Monthly Subscription");
                        bool bNoAnnualSKU = sku.Contains("Advanced Support for Partners Annual Subscription");
                        bool bMonthly = sku.Contains("No items were selected");
                        if (!bNoSKU && !bMonthly && !bNoAnnualSKU)
                        {
                            textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>{0}</span></li>", sku);
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
                        textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>{0}</span></li>", sku);
                    }
                }
            }

            textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
            if (data.PartnerArea != "")
            {
                textSummary.AppendLine(Resources.Submit_PreferredSAM_Location + " {0}", data.PartnerArea);
            }
            else
            {
                textSummary.AppendLine(Resources.Submit_PreferredSAM_Location + Resources.Submit_NoPreferedLocationSpecified);
            }
            textSummary.AppendLine("</span></li>");

            if (!bPromo)
            {
                textSummary.AppendLine("</span></li>");
                textSummary.AppendLine("<li class='MsoNormal' style='color: #333333; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto; mso-list: l1 level1 lfo3'><span style = 'font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif'>");
                if (data.SpecialInstructions != "")
                {
                    textSummary.AppendLine(Resources.Submit_SpecialInstructions + " {0}", data.SpecialInstructions);
                }
                else
                {
                    textSummary.AppendLine(Resources.Submit_SpecialInstructions + Resources.Submit_NoSpecialInstructionsProvided);
                }
            }

            textSummary.AppendLine("</span></li>");
            textSummary.AppendLine("</ul>");
            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");

            if (bPromo)
            {
                textSummary.AppendLine(Resources.Submit_Contact_PA + "<br/></span>");
            }
            else
            {
                textSummary.AppendLine(Resources.Submit_Contact_ASfP + "<br/></span>");
            }

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine(Resources.Submit_ContactUs + "<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine(Resources.Submit_Closing_ThankYou + "<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><br/>");
            textSummary.AppendLine(Resources.Submit_Sincerely + "<br/></span>");

            textSummary.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>");
            textSummary.AppendLine(Resources.Submit_ClosingSignature + "<br/></span>");

            // These substitutions are common
            custom.Add("###NOTIFICATION_HYPERLINK###", string.Empty);
            custom.Add("###MAIN_CONTENT###", textSummary.ToString());
            custom.Add("###PRIVACY_LINK###", Resources.Privacy_Link);

            custom.Add("###TERMS_OF_USE###", Resources.TOU_Link);
            custom.Add("###TRADEMARKS###", Resources.TrademarksURL);
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

            message.Subject = string.Format(Resources.Submit_Subject);
            //message.ToRecipients.Add("chads@microsoft.com"); //TODO: for testing only
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

            if (!string.IsNullOrEmpty(data.AdditionalEmail))
            {
                string[] addr = data.AdditionalEmail.Split(';');
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

            message.SendAndSaveCopy(WellKnownFolderName.SentItems);
        }
    }
}