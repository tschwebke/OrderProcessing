using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        /// <summary>
        /// Threaded allows us to avoid collisions with the other RegSys checking.
        /// </summary>

        public void ActivationCompleted_DoWork(object sender, DoWorkEventArgs e)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(ServerName), credentials);
            tfs.EnsureAuthenticated();
            WorkItemStore workItemStore = new WorkItemStore(tfs);
            List<WorkItem> workitems = new List<WorkItem>();

            string wiql;
            if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
            {
                wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP' AND [System.WorkItemType] = 'SKU Purchase' AND [System.State] = 'Validated' AND [Microsoft.VSTS.Common.SubState] = 'Active' AND [Microsoft.Operations.EmailControl.InvitationReminder]='Automatic' AND [Microsoft.Operations.Partners.Nomination.OriginalSource]='Test: ASfP RegSys Form' ORDER BY [System.Id]");
            }
            else
            {
                wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP' AND [System.WorkItemType] = 'SKU Purchase' AND [System.State] = 'Validated' AND [Microsoft.VSTS.Common.SubState] = 'Active' AND [Microsoft.Operations.EmailControl.InvitationReminder]='Automatic' AND [Microsoft.Operations.Partners.Nomination.OriginalSource]<>'Test: ASfP RegSys Form' ORDER BY [System.Id]");
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
                    wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email failed";
                    wi.Save();
                    //Read the item and check for basic validation
                    var errors = new List<string>();
                    Activation act = new Activation();

                    if (wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.MPNID.HQ") == "") { errors.Add("Missing HQ MPN ID"); } else { act.HQ_MPNID = wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.MPNID.HQ"); }
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Country.GMOBI.HQ") == "") { errors.Add("Missing HQ Country"); } else { act.HQ_Country = wi.GetFieldValue("Microsoft.Operations.Partners.Country.GMOBI.HQ"); }

                    act.FormLanguage = wi.GetFieldValue("Microsoft.Operations.Partners.FormLanguage");
                    act.PartnerOrg = wi.GetFieldValue("Microsoft.Operations.Partners.OrganizationName");
                    act.PartnerName = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Name");
                    act.VATIDValue = wi.GetFieldValue("Microsoft.Operations.Partners.ExternalIdentifiers.VATID");
                    act.SAMLead = wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.LeadName");
                    if (!wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.LeadEmail").IsValidEmailAddress() || wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.LeadEmail") == "") { errors.Add("SAM Lead Email is missing or malformed"); } else { act.SAMLeadEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.LeadEmail"); }

                    if (!wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Email").IsValidEmailAddress() || wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Email") == "") { errors.Add("Partner Contact Email is missing or malformed"); } else { act.PartnerContactEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Contact.Email"); }
                    act.PartnerBillingContact = wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactName");
                    act.PartnerBillingEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Billing.ContactEmail");
                    act.PartnerShipEmail = wi.GetFieldValue("Microsoft.Operations.Partners.ContactEmail.GMOBI");
                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate") == "") { errors.Add("Missing Servie Start Date"); } else { act.ServiceStartDate = Convert.ToDateTime(wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate")); }

                    if (wi.GetFieldValue("Microsoft.Operations.Partners.Partner.CurrentState") == "") { errors.Add("Missing Partner Current State"); } else { act.PartnerCurrentState = wi.GetFieldValue("Microsoft.Operations.Partners.Partner.CurrentState"); }

                    var noEmail = false;
                    switch (wi.GetFieldValue("Microsoft.Operations.Partners.Partner.CurrentState"))
                    {
                        case "PA Transition":
                        case "ASfP Net New Dynamics":

                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.Alias") == "") { errors.Add("Missing SAM Alias"); }
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.Alias") != "")
                            {
                                if (!wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.Alias").IsValidEmailAddress()) { errors.Add("SAM Email is malformed"); } else { act.SAMEmail = wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.Alias"); }
                            }
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.DisplayName") == "") { errors.Add("Missing SAM Name"); } else { act.SAM = wi.GetFieldValue("Microsoft.Operations.Partners.Support.SAM.DisplayName"); }
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount") == "") { errors.Add("Missing PA Custom Bill Amount"); } else { act.CustomBillAmount = wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount"); }
                            if (wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs") == "") { errors.Add("No SKUs were listed"); } else { act.SKUs = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs"); }
                            break;

                        case "ASfP CSP Accelerator":
                        case "ASfP Trial":
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate") == "") { errors.Add("Missing Servie Start Date"); } else { act.TrialServiceStartDate = Convert.ToDateTime(wi.GetFieldValue("Microsoft.Operations.Partners.Service.StartDate")); }
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.EndDate") == "") { errors.Add("Missing Service End Date"); } else { act.TrialServiceEndDate = Convert.ToDateTime(wi.GetFieldValue("Microsoft.Operations.Partners.Service.EndDate")); }
                            act.SKUs = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs");
                            break;

                        case "ASfP ISV":
                        case "ASfP Pilot":
                        case "ASfP Limited Preview":
                            noEmail = true;
                            break;

                        default: //Standard ASfP order
                            if (wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount") == "") { errors.Add("Missing Custom Bill Amount"); } else { act.CustomBillAmount = wi.GetFieldValue("Microsoft.Operations.Partners.Service.BillAmount"); }
                            if (wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs") == "") { errors.Add("No SKUs were listed"); } else { act.SKUs = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs"); }
                            break;
                    }
                    act.AssignedTo = wi.GetFieldValue("System.AssignedTo");

                    XDocument documentLang = XDocument.Load(string.Format(@"{0}\\Languages.xml", FileSystem.ExecutingFolder));
                    var displayLang =
                        (from c in documentLang.Root.Elements("LCID")
                         where (string)c.Attribute("language") == act.FormLanguage && c.Attribute("lang") != null
                         select c.Attribute("lang").Value).SingleOrDefault();
                    if (displayLang == "") { displayLang = "en"; }
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(displayLang);

                    //The Partner Current State selection is currently unsupported automation. ASfP ISV, ASfP Pilot, and ASfP Limited Preview
                    if (noEmail || errors.Count > 0)
                    {
                        if (noEmail)
                        {
                            try
                            {
                                string failureSubject = "Microsoft Advanced Support for Partners (ASfP) Activation Email Failure";
                                string errorName = "PartnerCurrentState";
                                string banner = "ASfP Activation Email Failure Report";

                                StringBuilder content = new StringBuilder();
                                var wiID = wi.GetFieldValue("System.Id");
                                content.AppendLine("There were errors in preparing the Activation Email for <a href='http://vstfpg07:8080/tfs/Operations/CSP/ASfP/_workItems#id={0}&triage=true&fullScreen=true&_a=edit'>workitem {0}</a><br/>", wiID);
                                content.AppendLine("<br/>");
                                content.AppendLine("Automated activation emails for {0} is currently not supported. The Automated Email delivery setting has been changed to Manual and ", wi.GetFieldValue("Microsoft.Operations.Partners.Partner.CurrentState"));
                                content.AppendLine("the Sub State is now Activation email failed.");
                                content.AppendLine("<br/>");

                                EmailErrors(content, failureSubject, errorName, banner, wi);
                            }
                            catch (Exception ex)
                            {
                                fails = 1;
                                wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email failed";
                                wi.History += string.Format("[SYSTEM] Tried to send activation failure email, but it failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                            }
                        }
                        else
                        {
                            try
                            {
                                string failureSubject = "Microsoft Advanced Support for Partners (ASfP) Activation Email Failure";
                                string errorName = "TFSProcessError";
                                string banner = "ASfP Activation Email Failure Report";

                                StringBuilder content = new StringBuilder();
                                var wiID = wi.GetFieldValue("System.Id");
                                content.AppendLine("There were errors in preparing the Activation Email for <a href='http://vstfpg07:8080/tfs/Operations/CSP/ASfP/_workItems#id={0}&triage=true&fullScreen=true&_a=edit'>workitem {0}</a><br/>", wiID);
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
                            }
                            catch (Exception ex)
                            {
                                fails = 1;
                                wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email failed";
                                wi.History += string.Format("[SYSTEM] Tried to send activation failure email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                            }
                        }
                        wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email failed";
                        wi.History += string.Format("[SYSTEM] Activation failed email was sent.");
                        wi.Save();
                    }
                    else
                    {
                        var activationType = "";
                        var frequency = "";

                        //TODO: Need to handle JA
                        List<string> order_lineitems = new List<string>(act.SKUs.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None));

                        switch (act.PartnerCurrentState)
                        {
                            case "PA Transition":
                            case "ASfP Net New Dynamics":
                                activationType = "PA";
                                if (order_lineitems.Count > 0)
                                {
                                    for (int i = 0; i < order_lineitems.Count; i++)
                                    {
                                        if (order_lineitems[i].Contains("PA Transition Package"))
                                        {
                                            order_lineitems.RemoveAt(i);
                                            i--;
                                        }
                                        else
                                        {
                                            var desc = order_lineitems[i].IndexOf("–", 10);
                                            order_lineitems[i] = order_lineitems[i].Substring(10, desc - 10).Trim();
                                        }
                                    }
                                }
                                break;

                            case "ASfP CSP Accelerator":
                            case "ASfP Trial":
                                activationType = "Trial";
                                break;

                            default:
                                activationType = "Standard";
                                if (order_lineitems.Count > 0)
                                {
                                    for (int i = 0; i < order_lineitems.Count; i++)
                                    {
                                        if (order_lineitems[i].Contains("Annual Subscription") || order_lineitems[i].Contains("Monthly Subscription") || order_lineitems[i].Contains("年間契約") || order_lineitems[i].Contains("月額課金"))
                                        {
                                            if (order_lineitems[i].Contains("Monthly Subscription") || order_lineitems[i].Contains("月額課金")) { frequency = "Monthly"; }
                                            else { frequency = "Annual"; }
                                            order_lineitems.RemoveAt(i);
                                            i--;
                                        }
                                        else
                                        {
                                            var desc = order_lineitems[i].IndexOf("–", 10);
                                            order_lineitems[i] = order_lineitems[i].Substring(10, desc - 10).Trim();
                                        }
                                    }
                                }
                                break;
                        }
                        //Advanced Support for Partners Annual Subscription – One - Time Payment, 3 % discount – USD 14, 550(local taxes apply)
                        //Advanced Support for Partners Monthly Subscription – Monthly Payment – USD 1, 250(local taxes apply)
                        //5 - Pack Non-Cloud Support Incidents – USD 1, 999(20 % discount)
                        //1 Cloud Consult Engagement – USD 2, 000(local taxes apply)
                        //1-Pack On Premises Support Incident – USD 499(local taxes apply)
                        //20 Hours of Services Account Management – USD 4, 000(local taxes apply)
                        //ASfP PA Transition Package – One - Time Payment
                        //ASfP Trial

                        Boolean sameEmail = false;
                        if (act.PartnerContactEmail == act.PartnerBillingEmail) { sameEmail = true; }

                        //VAT Requirement processing
                        XDocument document = XDocument.Load(string.Format(@"{0}\\VATID.xml", FileSystem.ExecutingFolder));
                        var displayText =
                            (from c in document.Root.Elements("country")
                             where (string)c.Attribute("name") == act.HQ_Country && c.Attribute("name") != null
                             select c.Attribute("Display").Value).SingleOrDefault();
                        Boolean vatID = false;

                        if (displayText != null)
                        {
                            if (act.VATIDValue == "")
                                vatID = true;
                        }
                        DateTime activationDate = DateTime.Today.SubscriptionFinalDate();
                        DateTime firstMondayDate = act.ServiceStartDate.StartOfWeek(DayOfWeek.Monday);

                        try
                        {
                            StringBuilder content = new StringBuilder();
                            Dictionary<string, string> custom = new Dictionary<string, string>();

                            //Salutation
                            content.AppendLine("<DIV style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Dear + "<br/>", act.PartnerName);
                            DateTime bizDays = DateTime.Today.AddBusinessDays(3);

                            //TODO: Should a SKU class be created and used here and in the Invoice Sending?
                            switch (activationType)
                            {
                                case "PA":
                                    switch (order_lineitems.Count)
                                    {
                                        case 0:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_1 + "</p>");
                                            break;

                                        case 1:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_2 + "</p>", order_lineitems[0].TrimStart());
                                            break;

                                        case 2:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_3 + "</p>", order_lineitems[0].TrimStart(), order_lineitems[1]);
                                            break;

                                        case 3:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_4 + "</p>", order_lineitems[0].TrimStart(), order_lineitems[1], order_lineitems[2]);
                                            break;

                                        case 4:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_5 + "</p>", order_lineitems[0], order_lineitems[1], order_lineitems[2], order_lineitems[3]);
                                            break;

                                        default:
                                            content.AppendLine("<p>" + Resources.Activate_PAIntro_1 + "</p>");
                                            break;
                                    }

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PA_Description + "</p>");

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PA_SAM_Description + "</p>", firstMondayDate.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)), act.SAM);

                                    if (sameEmail)
                                    {
                                        content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Email_Invoice, act.PartnerContactEmail.Trim(), act.CustomBillAmount);
                                    }
                                    else
                                    {
                                        content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Email_Invoice2, act.PartnerBillingContact, act.PartnerBillingEmail, act.CustomBillAmount);
                                    }

                                    content.AppendLine(Resources.Activate_PA_Active + "</p>");

                                    if (vatID && !sameEmail)
                                    {
                                        content.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_SendVAT_Blue + "</i></font></span></p>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));
                                    }

                                    if (sameEmail && vatID)
                                    {
                                        content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PleaseSend + "</p>");
                                        content.AppendLine("<ul type='disc'>");

                                        content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendVAT + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));

                                        content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendBillContact + "</li>");

                                        content.AppendLine("</ul>");
                                    }

                                    if (!vatID && sameEmail)
                                    {
                                        content.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_SendBillContact_Blue + "</i></font></span></p>");
                                    }

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Partnership + "</p>");
                                    break;

                                case "Trial":
                                    //TimeSpan trialLength = act.TrialServiceEndDate - act.TrialServiceStartDate;
                                    //content.AppendLine("<p>Thank you for choosing to participate in our {0}-day Trial for Advanced Support for Partners (ASfP).</p>", (double)trialLength.TotalDays);

                                    //content.AppendLine("<p>We have scheduled your ASfP trial activation for the week of {0:MMM dd, yyyy}. A confirmation email will be sent when", activationDate);
                                    //content.AppendLine("the configuration is complete.  Your {0}-day free trial will end {1:MMM dd, yyyy}. In order to avoid interruption to your service, ", (double)trialLength.TotalDays, act.TrialServiceEndDate);
                                    //content.AppendLine("please <a href='http://aka.ms/buyasfpnow'>subscribe to Advanced Support for Partners</a> before the end of your trial period. ");
                                    //content.AppendLine("If you have any questions, please contact your Services Account Manager (SAM) or email <a href='mailto:svcasfp@microsoft.com'>svcasfp@microsoft.com</a>.");
                                    //content.AppendLine("As soon as your account is configured, a Services Account Manager (SAM) will reach out and schedule a program overview session.</p>");

                                    //content.AppendLine("<p>In the meantime, if you have any questions or urgent cloud support requests please contact your Services Account Manager (SAM) or email <a href='mailto:svcasfp@microsoft.com'>svcasfp@microsoft.com</a>.</p>");
                                    wi.Fields["Microsoft.Operations.EmailControl.InvitationReminder"].Value = @"Manual";
                                    wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Active";
                                    wi.Save();
                                    break;

                                case "Standard":
                                    string standardSub;
                                    if (frequency == "Monthly") { standardSub = "<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Monthly; }
                                    else { standardSub = "<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Annual; }
                                    switch (order_lineitems.Count)
                                    {
                                        case 0:
                                            content.AppendLine("{0}.</p>", standardSub);
                                            break;

                                        case 1:
                                            content.AppendLine(Resources.Activate_Sub_Items_1 + "</p>", standardSub, order_lineitems[0].TrimStart());
                                            break;

                                        case 2:
                                            content.AppendLine(Resources.Activate_Sub_Items_2 + "</p>", standardSub, order_lineitems[0].TrimStart(), order_lineitems[1]);
                                            break;

                                        case 3:
                                            content.AppendLine(Resources.Activate_Sub_Items_3 + "</p>", standardSub, order_lineitems[0].TrimStart(), order_lineitems[1], order_lineitems[2]);
                                            break;

                                        case 4:
                                            content.AppendLine(Resources.Activate_Sub_Items_4 + "</p>", standardSub, order_lineitems[0].TrimStart(), order_lineitems[1], order_lineitems[2], order_lineitems[3]);
                                            break;

                                        default:
                                            content.AppendLine("{0}.</p>", standardSub);
                                            break;
                                    }

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_StartActivation + "</p>", activationDate.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PrgmOverview + "</p>");

                                    if (frequency == "Monthly")
                                    {
                                        if (!vatID && !sameEmail)
                                        {
                                            string billDate;
                                            if (DateTime.Today.Day >= 5 && DateTime.Today.Day <= 19) { billDate = "27"; } else { billDate = "15"; }
                                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_Invoicing_Date + "</i></font></span></p>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)), billDate);
                                        }
                                        if (vatID && !sameEmail)
                                        {
                                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PleaseSend + "</p>");
                                            content.AppendLine("<ul type='disc'>");

                                            string billDate;
                                            if (DateTime.Today.Day >= 5 && DateTime.Today.Day <= 19) { billDate = "27"; } else { billDate = "15"; }
                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Invoicing_Date + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)), billDate);
                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendVAT + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));

                                            content.AppendLine("</ul>");
                                        }

                                        if (sameEmail && vatID)
                                        {
                                            string billDate;
                                            if (DateTime.Today.Day >= 5 && DateTime.Today.Day <= 19) { billDate = "27"; } else { billDate = "15"; }
                                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PleaseSend + "</p>");
                                            content.AppendLine("<ul type='disc'>");

                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Invoicing_Date + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)), billDate);
                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendVAT + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));
                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendBillContact + "</li>");

                                            content.AppendLine("</ul>");
                                        }

                                        if (!vatID && sameEmail)
                                        {
                                            string billDate;
                                            if (DateTime.Today.Day >= 5 && DateTime.Today.Day <= 19) { billDate = "27"; } else { billDate = "15"; }
                                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PleaseSend + "</p>");
                                            content.AppendLine("<ul type='disc'>");

                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Invoicing_Date + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)), billDate);
                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendBillContact + "</li>");

                                            content.AppendLine("</ul>");
                                        }
                                    }
                                    else
                                    {
                                        if (vatID && !sameEmail)
                                        {
                                            content.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_SendVAT_Blue + "</i></font></span></p>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));
                                        }

                                        if (sameEmail && vatID)
                                        {
                                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_PleaseSend + "</p>");
                                            content.AppendLine("<ul type='disc'>");

                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_SendVAT + "</li>", bizDays.ToString("D", CultureInfo.CreateSpecificCulture(displayLang)));

                                            content.AppendLine("<li style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_SendBillContact + "</li>");

                                            content.AppendLine("</ul>");
                                        }

                                        if (!vatID && sameEmail)
                                        {
                                            content.AppendLine("<span style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'><font color='blue'><i>" + Resources.Activate_SendBillContact_Blue + "</i></font></span></p>");
                                        }
                                    }
                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_MeanTime + "</p>");

                                    content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_Partnership + "</p>");
                                    break;

                                default:
                                    wi.Fields["Microsoft.Operations.EmailControl.InvitationReminder"].Value = @"Manual";
                                    wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Active";
                                    wi.Save();
                                    break;
                            }
                            //Closing
                            content.AppendLine("<p style='font-size: 10.5pt; font-family: \"Segoe UI\",sans-serif; color: #333333'>" + Resources.Activate_ManyThanks + "</p><br/><br/></div>");

                            content.AppendLine("<div style='font-size:12.0pt;color:#848484'>" + Resources.Activate_Closing_Signature + "</div></DIV>");

                            custom.Add("###MAIN_CONTENT###", content.ToString());
                            custom.Add("###BANNER###", Resources.Activate_Banner);
                            custom.Add("###NOTIFICATION_HYPERLINK###", string.Empty);
                            custom.Add("###PRIVACY_LINK###", Resources.Privacy_Link);

                            custom.Add("###TERMS_OF_USE###", Resources.TOU_Link);
                            custom.Add("###TRADEMARKS###", Resources.TrademarksURL);
                            custom.Add("####.##.##.####", Assembly.GetExecutingAssembly().GetVersion());

                            string bodyText = FileSystem.GetEmbeddedFileContent("Microsoft.Operations.CSP.RegSys.Templates.Activation.html", "Microsoft.Operations.CSP.RegSys");
                            string compressionResult = string.Empty;

                            foreach (string s in custom.Keys)
                            {
                                bodyText = bodyText.Replace(s, custom[s]);
                            }

                            bodyText = Optimize.MinifyHtml(bodyText, out compressionResult, true);

                            ExchangeService svc = ExchangeServiceAutomatic.New(EmailUserIdentityEmail, EmailPassword);
                            EmailMessage message = new EmailMessage(svc);

                            if (act.FormLanguage == "Japanese")
                            {
                                message.Subject = string.Format(Resources.Activate_Subject);
                            }
                            else
                            {
                                message.Subject = string.Format(Resources.Activate_Subject, act.PartnerOrg);
                            }

                            message.ToRecipients.Add(act.PartnerContactEmail);
                            if (act.PartnerBillingEmail != "" && act.PartnerBillingEmail != null) { message.ToRecipients.Add(act.PartnerBillingEmail); }
                            if (act.PartnerShipEmail != "" && act.PartnerShipEmail != null) { message.ToRecipients.Add(act.PartnerShipEmail); }
                            if (act.SAMLeadEmail != "" && act.SAMLeadEmail != null) { message.CcRecipients.Add(act.SAMLeadEmail); }
                            if (act.SAMEmail != "" && act.SAMEmail != null) { message.CcRecipients.Add(act.SAMEmail); }
                            //message.CcRecipients.Add("chads@microsoft.com"); //TODO: for testing only
                            message.CcRecipients.Add("asfpsales@microsoft.com");
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

                            var fileNamePayload = Path.Combine(tempFileFolder, string.Format("activation_email.eml"));

                            using (var fileStream = new FileStream(fileNamePayload, FileMode.Create))
                            {
                                fileStream.Write(mimeContent.Content, 0, mimeContent.Content.Length);
                            }

                            message.SendAndSaveCopy(WellKnownFolderName.SentItems);
                            if (!string.IsNullOrEmpty(fileNamePayload))
                            {
                                try
                                {
                                    Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileNamePayload, "Activation Email");
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
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email failed";
                            wi.History += string.Format("[SYSTEM] There was an issue sending activation email: '<i>{0}</i>'", ex.Message);
                        }
                        var result = wi.Validate();

                        if (result.Count == 0 && fails == 0)
                        {
                            wi.Fields["Microsoft.Operations.Partners.Emails.ActivateDate"].Value = DateTime.Today;
                            wi.Fields["Microsoft.VSTS.Common.SubState"].Value = @"Activation email sent";
                            wi.History += "[SYSTEM] Activation email has been successfully sent.";
                        }
                        else
                        {
                            count_non_changed_items++;
                        }
                    }
                }
                else
                {
                    count_non_changed_items++;
                }
                //TODO: Add logic to add error if save is not possible
                //TODO: Add this to invoice sending also
                //foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result)
                //{
                //    Console.WriteLine(wi.Id + info.Status);
                //}
                try
                {
                    wi.Save();
                }
                catch (Exception ex)
                {
                    wi.History += string.Format("[SYSTEM] Failed to save work item: '<i>{0}</i>'", ex.Message);
                }
            }
        }

        public void ActivationCompleted_InvokeThread()
        {
            if (ActivationCompleted != null && !ActivationCompleted.IsBusy)
            {
                ActivationCompleted.RunWorkerAsync();
            }
        }
    }
}