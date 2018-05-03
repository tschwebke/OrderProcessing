using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        /// <summary>
        /// Threaded allows us to avoid collisions with the other RegSys checking. This is because
        /// the invoice processing may take some time!
        /// </summary>

        public void InvoiceSending_DoWork(object sender, DoWorkEventArgs e)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(ServerName), credentials);
            tfs.EnsureAuthenticated();
            WorkItemStore workItemStore = new WorkItemStore(tfs);
            List<WorkItem> workitems = new List<WorkItem>();

            string wiql = "";
            if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "thads")
            {
                wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP'  AND  [System.WorkItemType] = 'SKU Purchase' AND  [System.State] = 'Entitled'  AND  [Microsoft.VSTS.Common.SubState] = 'Release for Invoice' AND [Microsoft.Operations.EmailControl.InvitationReminder]='Automatic' AND [Microsoft.Operations.Partners.Nomination.OriginalSource]='Test: ASfP RegSys Form' ORDER BY [System.Id]");
            }
            else
            {
                wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP'  AND  [System.WorkItemType] = 'SKU Purchase' AND  [System.State] = 'Entitled'  AND  [Microsoft.VSTS.Common.SubState] = 'Release for Invoice' AND [Microsoft.Operations.EmailControl.InvitationReminder]='Automatic' AND [Microsoft.Operations.Partners.Nomination.OriginalSource]<>'Test: ASfP RegSys Form' ORDER BY [System.Id]");
            }

            workitems = workItemStore.ExecuteQueryText(wiql);

            int count_non_changed_items = 0;
            var mainSKU = "";
            var fails = 0;
            foreach (WorkItem wi in workitems)
            {
                if (wi.LastChangedAgeInMinutes(1))
                {
                    wi.Open();
                    wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Failed to Invoice";
                    wi.Save();
                    //Read the item and check for basic validation
                    var errors = new List<string>();
                    Invoice inv = new Invoice();
                    if (wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs") == "") { errors.Add("No SKUs were listed"); }
                    else
                    {
                        inv.SKUs_OriginalInformation = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs");
                        if (inv.SKUs_OriginalInformation.Contains("PA Transition Package"))
                        {
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount") == "") { errors.Add("Bill Amount is required for PA Transition order"); } else { inv.BillAmount = wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount"); }
                            if (wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.BillingFrequency") == "") { errors.Add("Bill Frequency is required for PA Transition order"); } else { inv.BillFrequency = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.BillingFrequency"); }
                        }
                        else
                        {
                            inv.BillAmount = wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount");
                            inv.BillFrequency = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.BillingFrequency");
                        }
                    }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.OrganizationName") == "") { errors.Add("Missing Partner Organization Name"); } else { inv.PartnerOrg = wi.GetFieldValue("Microsoft.Operations.Partners.OrganizationName"); }
                    inv.IsSAP = wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.SAP.IsKnown");

                    //Sold to
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Name") == "") { errors.Add("Missing Sold to Contact Name"); } else { inv.PartnerContactName = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Name"); }
                    inv.PartnerContactEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Email");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Address1") == "") { errors.Add("Missing Sold to Address 1"); } else { inv.PartnerAddressLine1 = wi.GetFieldValue("Microsoft.Operations.Partners.Address1"); }
                    inv.PartnerAddressLine2 = wi.GetFieldValue("Microsoft.Operations.Partners.Address2");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Postcode") == "") { errors.Add("Missing Sold to Zip Code"); } else { inv.PartnerZip = wi.GetFieldValue("Microsoft.Operations.Partners.Postcode"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.City") == "") { errors.Add("Missing Sold to City"); } else { inv.PartnerCity = wi.GetFieldValue("Microsoft.Operations.Partners.City"); }
                    inv.PartnerState = wi.GetFieldValue("Microsoft.Operations.Partners.Province");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Country") == "") { errors.Add("Missing Sold to Country"); } else { inv.PartnerCountry = wi.GetFieldValue("Microsoft.Operations.Partners.Country"); }

                    //Bill to
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactName") == "") { errors.Add("Missing Billing Contact Name"); } else { inv.PartnerBillingContactName = wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactName"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactEmail") == "") { errors.Add("Missing Billing Contact Email"); } else { inv.PartnerBillingEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactEmail"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Address1.Billing") == "") { errors.Add("Missing Billing to Address 1"); } else { inv.PartnerBillingAddressLine1 = wi.GetFieldValue("Microsoft.Operations.Partners.Address1.Billing"); }
                    inv.PartnerBillingAddressLine2 = wi.GetFieldValue("Microsoft.Operations.Partners.Address2.Billing");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Postcode.Billing") == "") { errors.Add("Missing Billing to Zip Code"); } else { inv.PartnerBillingZip = wi.GetFieldValue("Microsoft.Operations.Partners.Postcode.Billing"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.City.Billing") == "") { errors.Add("Missing Billing to City"); } else { inv.PartnerBillingCity = wi.GetFieldValue("Microsoft.Operations.Partners.City.Billing"); }
                    inv.PartnerBillingState = wi.GetFieldValue("Microsoft.Operations.Partners.Province.Billing");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Country.Billing") == "") { errors.Add("Missing Billing to Country"); } else { inv.PartnerBillingCountry = wi.GetFieldValue("Microsoft.Operations.Partners.Country.Billing"); }

                    //Ship to
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.ContactName.GMOBI") == "") { errors.Add("Missing Ship to Contact Name"); } else { inv.PartnerShipContactName = wi.GetFieldValue("Microsoft.Operations.Partners.ContactName.GMOBI"); }
                    inv.PartnerShipEmail = wi.GetFieldValue("Microsoft.Operations.Partners.ContactEmail.GMOBI");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Address1.GMOBI") == "") { errors.Add("Missing Ship to Address 1"); } else { inv.PartnerShipAddressLine1 = wi.GetFieldValue("Microsoft.Operations.Partners.Address1.GMOBI"); }
                    inv.PartnerShipAddressLine2 = wi.GetFieldValue("Microsoft.Operations.Partners.Address2.GMOBI");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Postcode.GMOBI") == "") { errors.Add("Missing Ship to Zip Code"); } else { inv.PartnerShipZip = wi.GetFieldValue("Microsoft.Operations.Partners.Postcode.GMOBI"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.City.GMOBI") == "") { errors.Add("Missing Ship to City"); } else { inv.PartnerShipCity = wi.GetFieldValue("Microsoft.Operations.Partners.City.GMOBI"); }
                    inv.PartnerShipState = wi.GetFieldValue("Microsoft.Operations.Partners.Province.Billing");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Country.GMOBI") == "") { errors.Add("Missing Ship to Country"); } else { inv.PartnerShipCountry = wi.GetFieldValue("Microsoft.Operations.Partners.Country.GMOBI"); }

                    inv.Tax_VATID = wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.VATID");
                    if (wi.GetFieldValue("Microsoft.Finance.Purchase.Invoice.PaymentTerms.Standard") == "") { errors.Add("Missing Payment Terms"); } else { inv.PaymentTerms = wi.GetFieldValue("Microsoft.Finance.Purchase.Invoice.PaymentTerms.Standard"); }
                    inv.PaymentTerms = "Net " + inv.PaymentTerms + " days";
                    if (wi.GetFieldValue("Microsoft.Finance.Purchase.Currency.Billing") == "") { errors.Add("Missing Currency"); } else { inv.Currency = wi.GetFieldValue("Microsoft.Finance.Purchase.Currency.Billing"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Tax.IsExempt") == "") { errors.Add("Is Exempt Flag not set"); } else { inv.TaxStatus = wi.GetFieldValue("Microsoft.Operations.Partners.Tax.IsExempt"); }
                    if (wi.GetFieldValue("Microsoft.Finance.Purchase.Invoice.DeliveryMethod") == "") { errors.Add("Missing Invoice Delivery Method"); } else { inv.InvoiceDeliveryMethod = wi.GetFieldValue("Microsoft.Finance.Purchase.Invoice.DeliveryMethod"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate") == "") { errors.Add("Missing Service Start Date"); } else { inv.ServiceStartDate = wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.EndDate") == "") { errors.Add("Missing Service End Date"); } else { inv.ServiceEndDate = wi.GetFieldValue("Microsoft.Operations.Partners.Service.EndDate"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.FirstBillDate") == "") { errors.Add("Missing Desired Bill Date"); } else { inv.DesireBillDate = wi.GetFieldValue("Microsoft.Operations.Partners.Service.FirstBillDate"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.FirstBill.Date") == "") { errors.Add("Missing First Bill Date"); } else { inv.FirstBillDate = wi.GetFieldValue("Microsoft.Operations.Partners.Service.FirstBill.Date"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.CMATDomains") == "") { errors.Add("Missing CMAT Domains"); } else { inv.CMATDomain = wi.GetFieldValue("Microsoft.Operations.Partners.CMATDomains"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.MPNID.HQ") == "") { errors.Add("Missing HQ MPN ID"); } else { inv.HQMPNID = wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.MPNID.HQ"); }

                    inv.AssignedTo = wi.GetFieldValue("System.AssignedTo");

                    if (errors.Count > 0)
                    {
                        try
                        {
                            string failureSubject = "Microsoft Advanced Support for Partners (ASfP) Invoice Failure";
                            string errorName = "InvoicingError";
                            string banner = "ASfP Manual Invoice Request Failure Report";

                            StringBuilder content = new StringBuilder();
                            var wiID = wi.GetFieldValue("System.Id");
                            content.AppendLine("There were errors in preparing the XLS to send to the GOC for <a href='http://vstfpg07:8080/tfs/Operations/CSP/ASfP/_workItems#id={0}&triage=true&fullScreen=true&_a=edit'>workitem {0}</a><br/>", wiID);
                            content.AppendLine("<br/>");
                            content.AppendLine("Below is a list of the errors found:<br/>");
                            content.AppendLine("<ul>");
                            foreach (var err in errors)
                            {
                                if (!err.IsBlank()) { content.AppendLine("<li>{0}</li>", err); }
                            }
                            content.AppendLine("</ul>");
                            content.AppendLine("<br/>");

                            EmailErrors(content, failureSubject, errorName, banner, wi);

                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Failed to Invoice";
                            wi.Save();
                        }
                        catch (Exception ex)
                        {
                            fails = 1;
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Failed to Invoice";
                            wi.History += string.Format("[SYSTEM] Tried to send invoice failed email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                        }
                    }
                    else
                    {
                        inv.ServiceEndDate = inv.ServiceEndDate.Substring(0, inv.ServiceEndDate.IndexOf(" "));
                        inv.ServiceStartDate = inv.ServiceStartDate.Substring(0, inv.ServiceStartDate.IndexOf(" "));
                        inv.FirstBillDate = inv.FirstBillDate.Substring(0, inv.FirstBillDate.IndexOf(" "));
                        string tempFilename = Path.Combine(Path.GetTempPath(), string.Format("{0}.xlsx", Guid.NewGuid().ToString()));
                        //Create spreadsheet dynamically using all the content from this record.
                        try
                        {
                            if (File.Exists(tempFilename)) File.Delete(tempFilename);
                            using (FileStream fs = File.Create(tempFilename))
                            {
                                byte[] bytes = Assembly.GetExecutingAssembly().GetManifestResourceStream("Microsoft.Operations.CSP.RegSys.Templates.CSP AfSP Manual Billing Form v1 01 22 16.xlsx").ReadFully();
                                fs.Write(bytes, 0, bytes.Length);
                            }
                            ExcelPackage pck = new ExcelPackage(new FileInfo(tempFilename));
                            ExcelWorksheet ws = pck.Workbook.Worksheets["Customer Information"];

                            ws.Cells[2, 2].Value = inv.PartnerOrg;
                            ws.Cells[3, 2].Value = inv.IsSAP;
                            ws.Cells[4, 2].Value = inv.AssignedTo;
                            ws.Cells[5, 2].Value = DateTime.Today.ToString("d");

                            //Sale to
                            ws.Cells[9, 2].Value = inv.PartnerContactName;
                            ws.Cells[10, 2].Value = inv.PartnerAddressLine1;
                            ws.Cells[11, 2].Value = inv.PartnerAddressLine2;
                            ws.Cells[13, 2].Value = inv.PartnerZip;
                            ws.Cells[14, 2].Value = inv.PartnerCity;
                            ws.Cells[15, 2].Value = inv.PartnerState;
                            ws.Cells[16, 2].Value = inv.PartnerCountry;

                            //Bill to
                            ws.Cells[18, 2].Value = inv.PartnerBillingContactName;
                            ws.Cells[19, 2].Value = inv.PartnerBillingAddressLine1;
                            ws.Cells[20, 2].Value = inv.PartnerBillingAddressLine2;
                            ws.Cells[22, 2].Value = inv.PartnerBillingZip;
                            ws.Cells[23, 2].Value = inv.PartnerBillingCity;
                            ws.Cells[24, 2].Value = inv.PartnerBillingState;
                            ws.Cells[25, 2].Value = inv.PartnerBillingCountry;

                            //Ship to
                            ws.Cells[27, 2].Value = inv.PartnerShipContactName;
                            ws.Cells[28, 2].Value = inv.PartnerShipAddressLine1;
                            ws.Cells[29, 2].Value = inv.PartnerShipAddressLine2;
                            ws.Cells[31, 2].Value = inv.PartnerShipZip;
                            ws.Cells[32, 2].Value = inv.PartnerShipCity;
                            ws.Cells[33, 2].Value = inv.PartnerShipState;
                            ws.Cells[34, 2].Value = inv.PartnerShipCountry;

                            ws.Cells[35, 2].Value = inv.Tax_VATID;
                            ws.Cells[36, 2].Value = inv.PaymentTerms;
                            ws.Cells[37, 2].Value = inv.Currency;
                            ws.Cells[38, 2].Value = inv.TaxStatus;
                            ws.Cells[40, 2].Value = inv.PartnerBillingEmail;
                            ws.Cells[41, 2].Value = inv.InvoiceDeliveryMethod;
                            ws.Cells[42, 2].Value = wi.Id;
                            ws.Cells[43, 2].Value = inv.HQMPNID;
                            ws.Cells[44, 2].Value = inv.CMATDomain;

                            ws = pck.Workbook.Worksheets["Billing Data"];
                            ws.Cells[3, 2].Value = inv.PartnerOrg;

                            string existingSkuInformation = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs");
                            List<string> order_lineitems = new List<string>(existingSkuInformation.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None));

                            int row = 10;
                            foreach (var item in order_lineitems)
                            {
                                ws.Cells[row, 2].Value = item.Substring(0, 9);
                                ws.Cells[row, 3].Value = "1";
                                ws.Cells[row, 4].Style.Numberformat.Format = "@";
                                ws.Cells[row, 6].Style.Numberformat.Format = "@";
                                ws.Cells[row, 6].Value = inv.ServiceStartDate;
                                ws.Cells[row, 7].Style.Numberformat.Format = "@";
                                ws.Cells[row, 7].Value = inv.ServiceEndDate;
                                ws.Cells[row, 8].Value = inv.DesireBillDate;

                                if (item.Contains("PA Transition Package") || item.Contains("[SYSTEM] No SKU items"))
                                {
                                    ws.Cells[row, 10].Value = "ASfP PA Transition Package";
                                    ws.Cells[row, 4].Value = inv.BillAmount;
                                    ws.Cells[row, 9].Value = inv.BillFrequency;
                                }
                                else
                                {
                                    var desc = item.IndexOf("–", 10);
                                    ws.Cells[row, 10].Value = item.Substring(10, desc - 10).Trim();

                                    if (inv.Currency == "USD")
                                    {
                                        var amt = item.IndexOf("USD") + 4;
                                        ws.Cells[row, 4].Value = item.Substring(amt, item.IndexOf("(") - amt).Trim();
                                    }
                                    if (inv.Currency == "JPY" || inv.Currency == "JPN")
                                    {
                                        if (item.Contains("Advanced Support for Partners"))
                                        {
                                            var amt = item.IndexOf("–", 46) + 2;
                                            ws.Cells[row, 4].Value = item.Substring(amt, item.IndexOf("円") - amt).Trim();
                                        }
                                        else
                                        {
                                            var amt = item.IndexOf("–") + 2;
                                            ws.Cells[row, 4].Value = item.Substring(amt, item.IndexOf("円") - amt).Trim();
                                        }
                                    }

                                    if (item.Contains("Advanced Support for Partners"))
                                    {
                                        if (item.Contains("Monthly") || item.Contains("月額課金")) { ws.Cells[row, 9].Value = "Monthly"; } else { ws.Cells[row, 9].Value = "One Time"; }
                                    }
                                    else
                                    {
                                        ws.Cells[row, 9].Value = "One Time";
                                    }
                                }
                                row = row + 1;
                            }
                            pck.Save();
                            Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment original_xlsx = new TeamFoundation.WorkItemTracking.Client.Attachment(tempFilename, inv.PartnerOrg + " Invoice");
                            wi.Attachments.Add(original_xlsx);
                        }
                        catch (Exception ex)
                        {
                            fails = 1;
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Failed to Invoice";
                            wi.History += string.Format("[SYSTEM] There was an issue creating the invoice spreadsheet: '<i>{0}</i>'", ex.Message);
                        }

                        try
                        {
                            //var HQ_Country=wi.GetFieldValue("Microsoft.Operations.Partners.Country.GMOBI.HQ");
                            //XDocument document = XDocument.Load(string.Format(@"{0}\\VATID.xml", FileSystem.ExecutingFolder));
                            //var displayText =
                            //    (from c in document.Root.Elements("country")
                            //     where (string)c.Attribute("name") == HQ_Country && c.Attribute("name") != null
                            //     select c.Attribute("Display").Value).SingleOrDefault();

                            StringBuilder content = new StringBuilder();
                            Dictionary<string, string> custom = new Dictionary<string, string>();

                            content.AppendLine("Please create an invoice for {0}<br/>", inv.PartnerOrg);

                            //if (displayText != null)
                            //{
                            //    content.AppendLine("<p><font color='red'><i>VAT ID requirement: {0}</i></font></p>", displayText);
                            //}

                            content.AppendLine("<ul>");

                            if (inv.SKUs_OriginalInformation.Contains("PA Transition Package")) { mainSKU = "ASfP PA Transition Package"; } else { mainSKU = "Advanced Support for Partners Subscription"; }

                            content.AppendLine("<li>Billing period covers {0} - {1} for {2}</li>", inv.ServiceStartDate, inv.ServiceEndDate, mainSKU);
                            if (inv.SKUs_OriginalInformation.Contains("Monthly"))
                            {
                                content.AppendLine("<li>Regular bill date is {0}</li>", inv.DesireBillDate);
                            }
                            else
                            {
                                content.AppendLine("<li>Regular bill date is {0}</li>", inv.FirstBillDate);
                            }
                            content.AppendLine("</ul>");

                            content.AppendLine("See attached order form for details. If you have questions regarding this order form, please contact <a href='mailto:Svcasfp@microsoft.com'>Microsoft Advanced Support for Partners</a>");
                            content.AppendLine("<br/>");
                            content.AppendLine("<br/>");
                            content.AppendLine("Thanks,");
                            content.AppendLine("<br/>");
                            content.AppendLine("Advanced Support for Partners Team");

                            custom.Add("###MAIN_CONTENT###", content.ToString());
                            custom.Add("###BANNER###", "ASfP Manual Invoice Request");
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

                            bodyText = Optimize.MinifyHtml(bodyText, out compressionResult, true);

                            ExchangeService svc = ExchangeServiceAutomatic.New(EmailUserIdentityEmail, EmailPassword);
                            EmailMessage message = new EmailMessage(svc);

                            message.Subject = string.Format("Manual SAP Invoice Order for ASfP; {0}", inv.PartnerOrg);
                            message.ToRecipients.Add("Svcasfp@microsoft.com");
                            message.ReplyTo.Add("svcasfp@microsoft.com");
                            message.Body = new MessageBody(BodyType.HTML, bodyText);
                            Email.InsertImageFromResource(ref message, "microsoft_footer.png");
                            Email.InsertImageFromResource(ref message, "asfp_banner.png");
                            byte[] b = File.ReadAllBytes(tempFilename);

                            message.Attachments.AddFileAttachment(string.Format("{0} Invoice Order {1}.xlsx", inv.PartnerOrg, DateTime.Today.ToString("yyyyMMdd")), b);
                            message.Save(WellKnownFolderName.Drafts); // this is required to get the "ID" value so we can access other properties of the object. After the mail gets sent, this is removed from DRAFTS

                            message.Load(new PropertySet(ItemSchema.MimeContent));
                            var mimeContent = message.MimeContent;

                            string tempFileFolder = string.Format(@"{0}\{1}\Temp", FileSystem.BaseFolder, "CSP Workflow");
                            if (!Directory.Exists(tempFileFolder)) Directory.CreateDirectory(tempFileFolder);

                            var fileNamePayload = Path.Combine(tempFileFolder, string.Format("invoice_email.eml"));

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
                            fails = 1;
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Release for Invoice";
                            wi.History += string.Format("[SYSTEM] There was an issue sending invoice spreadsheet via email to the GOC: '<i>{0}</i>'", ex.Message);
                        }
                        //Attach back a copy of the invoice to this work item so that we have a complete audit trail of what was sent (no need to record the email body text)
                        var result = wi.Validate();
                        foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result)
                        {
                            Console.WriteLine(wi.Id + info.Status);
                        }

                        if (result.Count == 0 && fails == 0)
                        {
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Sent for Invoice";
                            wi.Fields["System.AssignedTo"].Value = @"[CSP]\GOC";
                            wi.History += "[SYSTEM] Invoice has been sent to GOC. Now awaiting GOC processing and approval.";
                        }
                        else
                        {
                            count_non_changed_items++;
                        }
                        try
                        {
                            wi.Save();
                        }
                        catch (Exception ex)
                        {
                            wi.History += string.Format("[SYSTEM] Failed to save work item: '<i>{0}</i>'", ex.Message);
                        }
                        // CLEANUP THE FILE SYSTEM NOW WE'RE FINISHED
                        if (File.Exists(tempFilename)) File.Delete(tempFilename);
                    }
                }
                else
                {
                    count_non_changed_items++;
                }
                wi.Save();
            }
        }

        public void InvoiceSending_InvokeThread()
        {
            if (InvoiceSending != null && !InvoiceSending.IsBusy)
            {
                InvoiceSending.RunWorkerAsync();
            }
        }
    }
}