using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        private Dictionary<string, int> ExistingMatchesAlreadyKnown = new Dictionary<string, int>();

        public async void ProcessWizardEntries_DoWork(object sender, DoWorkEventArgs e)
        {
            Collection<RegSysWizardItem> recent_submissions = new Collection<RegSysWizardItem>();

            DateTime today = DateTime.Today.AddDays(1); // go a day ahead to avoid any international time conflicts

            string Url = "";
            if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
            {
                DateTime pToday = DateTime.Today.AddDays(-5); // and a day backward - note: if it is known that some days were skipped, please adjust
                Url = "http://co1msftolappa02/RegSysReportsApi/api/WizardResponse?wizardID=84dd6e58-e97a-4091-b632-bb0df8eb4a48&fromDate=" + pToday.ToString("yyyy-MM-dd") + "&toDate=" + today.ToString("yyyy-MM-dd");
            }
            else
            {
                DateTime pToday = DateTime.Today.AddDays(-7); // and a day backward - note: if it is known that some days were skipped, please adjust
                Url = "http://co1msftolappa02/RegSysReportsApi/api/WizardResponse?wizardID=8d076345-5e8c-4e3f-9d7a-1fc3c9be69fd&fromDate=" + pToday.ToString("yyyy-MM-dd") + "&toDate=" + today.ToString("yyyy-MM-dd");
            }

            XmlDocument xmlDoc = new XmlDocument();
            WebRequest webRequest = WebRequest.Create(Url);
            webRequest.ContentType = "text/xml";
            Uri uri = new Uri("http://co1msftolappa02/");

            // IMPORTANT: This web request needs to be in the security context of the 'ASfP
            //            Automation' account, otherwise it won't work. i.e. the RegSys 'report' has
            // very specific (and unique) view requirements. You'll also need to use this account to
            // do any casual viewing/testing. Password change history and documentation exists at:
            // $/Online Operations/Infrastructure/Documentation/Maintenance and Operation/Passwords
            // for 'ASfP Automation'.txt

            webRequest.Credentials = credentials;

            using (WebResponse webResponse = await webRequest.GetResponseAsync())
            {
                using (Stream responseStream = webResponse.GetResponseStream())
                {
                    xmlDoc.Load(responseStream);
                }
            }

            var mgr = new XmlNamespaceManager(xmlDoc.NameTable);
            mgr.AddNamespace("b", "http://schemas.datacontract.org/2004/07/RegSysReports.Models");
            mgr.AddNamespace("d4p1", "http://schemas.microsoft.com/2003/10/Serialization/Arrays");

            // The conversion of the raw data to a RegSysWizardItem has some unique and subtle
            // considerations to it. If you are working extensively with the object, be sure to view
            // the notes (inside 'RegSysWizardItem.cs')

            foreach (XmlElement x in xmlDoc.SelectNodes("//b:WizardResponse", mgr))
            {
                recent_submissions.Add(new RegSysWizardItem(x));
            }

            // Now, get that listing and compare against ones which have already been created in the
            // system. We will have to perform an individual lookup for each, but because we don't do
            // it often, this should be fine.

            // If we don't have a record for it, then we should process it. This will be based on the
            // value from 'VirtualKey', which will be stored in TFS

            TfsTeamProjectCollection tfs = null;
            WorkItemStore workItemStore = null;
            bool tfsPreconnectionRequired = false; // assume not

            foreach (RegSysWizardItem data in recent_submissions)
            {
                if (!ExistingMatchesAlreadyKnown.ContainsKey(data.VirtualKey))
                {
                    tfsPreconnectionRequired = true;
                    break;
                }
            }

            if (tfsPreconnectionRequired)
            {
                tfs = new TfsTeamProjectCollection(new Uri(ServerName), credentials);
                tfs.EnsureAuthenticated();
                workItemStore = new WorkItemStore(tfs);
            }

            foreach (RegSysWizardItem data in recent_submissions)
            {
                // To save some work, we'll check our in-memory record of items that we've already
                // looked at. If we know that an item already exists then we don't have to do use the
                // TFS Lookup, which means we can handle more volume, smarter.

                if (ExistingMatchesAlreadyKnown.ContainsKey(data.VirtualKey))
                {
                    // If this is found, then it means we already know (within this session) that the
                    // item has been processed in TFS and there is no need for any further work.
                }
                else
                {
                    // Perform the work to copy this to TFS, which first involves running a check to
                    // see if it does actually exist in TFS.
                    string wiql = "";
                    if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
                    {
                        wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP'  AND  [System.WorkItemType] = 'SKU Purchase'  AND  [Microsoft.Bios.WorkItem.ID.ExternalReference.ClientNumber] = '{0}' AND [Microsoft.Operations.Partners.Nomination.OriginalSource]='Test: ASfP RegSys Form' ORDER BY [System.Id] ", data.VirtualKey);
                    }
                    else
                    {
                        wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP'  AND  [System.WorkItemType] = 'SKU Purchase'  AND  [Microsoft.Bios.WorkItem.ID.ExternalReference.ClientNumber] = '{0}' ORDER BY [System.Id] ", data.VirtualKey);
                    }

                    List<WorkItem> existing_matches = workItemStore.ExecuteQueryText(wiql);

                    if (existing_matches.Count >= 1)
                    {
                        // According to the query, there is already at least one entry which matches
                        // this RegSys record (based on our Virtual Key). We don't need to take any
                        // further action, and we will likewise we will make a note of the match so
                        // that we don't need to check it again. In the extremely unlikely event of
                        // more than one match, we'll only take the first entry and ignore the rest.

                        data.ExistingTfsID = existing_matches[0].Id;
                        try
                        {
                            ExistingMatchesAlreadyKnown.Add(data.VirtualKey, data.ExistingTfsID);
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    else if (existing_matches.Count == 0)
                    {
                        // An existing entry for this does NOT already exist, therefore we can go
                        // ahead and *create it*.
                        WorkItemType template = workItemStore.Projects["CSP"].WorkItemTypes["SKU Purchase"];
                        WorkItem w = template.NewWorkItem();

                        w.Fields["Microsoft.Bios.WorkItem.ID.ExternalReference.ClientNumber"].Value = data.VirtualKey;
                        XDocument document = XDocument.Load(string.Format(@"{0}\\Languages.xml", FileSystem.ExecutingFolder));
                        var displayLanguage =
                            (from c in document.Root.Elements("LCID")
                             where (string)c.Attribute("id") == data.LCID && c.Attribute("language") != null
                             select c.Attribute("language").Value).SingleOrDefault();
                        var displayLang =
                            (from c in document.Root.Elements("LCID")
                             where (string)c.Attribute("id") == data.LCID && c.Attribute("lang") != null
                             select c.Attribute("lang").Value).SingleOrDefault();
                        ResourceHelper res = new ResourceHelper("Microsoft.Operations.CSP.RegSys.Resources", GetType().Assembly);

                        //TODO: For testing only
                        //if (displayLang == "en") { continue; }
                        var resValue = "";
                        if (data.PartnerArea != "")
                        {
                            var resKey = res.GetResourceName(data.PartnerArea.Trim(), displayLang);
                            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en");
                            resValue = res.GetResourceValue(resKey);
                        }

                        w.Fields["Microsoft.Operations.Partners.Support.GeographicalPreference"].Value = resValue;
                        w.Fields["Microsoft.Operations.Partners.FormLanguage"].Value = displayLanguage;
                        w.Fields["Microsoft.Operations.Partners.Contact.Name"].Value = data.PartnerContactName.Trim();
                        w.Fields["Microsoft.Operations.Partners.Contact.Email"].Value = data.PartnerContactEmail.Trim();

                        if (!data.SpecialInstructions.Trim().Contains("PATN"))
                        {
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Request.Notes"].Value = data.SpecialInstructions.Trim();
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Request.SKUs"].Value = data.SKUs;
                            w.Fields["Microsoft.Operations.Partners.Partner.CurrentState"].Value = @"ASfP Purchase";
                        }
                        else
                        {
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Request.Notes"].Value = data.SpecialInstructions.Trim();
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Request.SKUs"].Value = data.SKUs;
                            w.Fields["Microsoft.Operations.Partners.Service.PromoCode"].Value = data.SpecialInstructions.Trim();
                            w.Fields["Microsoft.Operations.Partners.Partner.CurrentState"].Value = @"PA Transition";
                        }

                        w.Fields["Microsoft.Operations.Partners.Billing.ContactName"].Value = data.BillingContactName.Trim();
                        w.Fields["Microsoft.Operations.Partners.Billing.ContactEmail"].Value = data.BillingContactEmail.Trim();
                        w.Fields["Microsoft.Operations.Partners.AreaOfExpertise"].Value = resValue;
                        w.Fields["Microsoft.Operations.Partners.ExternalIdentifiers.MPNID"].Value = data.MPNID.Trim();
                        w.Fields["Microsoft.Operations.CreditDiscountApproval.Company.TenantDomain"].Value = data.CustomerTenantDomain.Trim();
                        w.Fields["Microsoft.Operations.AutomatedProcess.LastChecked"].Value = data.CreatedDateInRegSys; // store here for now ...

                        if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
                        {
                            w.Fields["Microsoft.Operations.Partners.Nomination.OriginalSource"].Value = "Test: " + data.DataOrigin;
                        }
                        else
                        {
                            w.Fields["Microsoft.Operations.Partners.Nomination.OriginalSource"].Value = data.DataOrigin;
                        }
                        //w.Fields["Microsoft.Operations.WorkItem.Ownership.RequestOrigin"].Value = data.DataOriginDetail;
                        w.Fields["System.AssignedTo"].Value = @""; // default

                        w.History += string.Format("<br/>Imported from RegSys: [{0}]<br/>", data.VirtualKey);
                        w.History += "Email Verified? " + data.IsEmailVerified + "<br/>";
                        w.History += "IP Address: " + data.IPAddress + "<br/>";
                        w.History += "Agreement Recorded: '<i>" + data.AgreementToTerms + "</i>'" + "<br/><br/>";

                        w.CleanSkuInformation(true); // Added support - SKU Number insertion, refer to Task 67509 for details and maintenance.

                        foreach (string warning in data.Warnings)
                        {
                            w.History += "[WARNING] " + warning + "<br/>";
                        }

                        var result = w.Validate();
                        foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result)
                        {
                            data.Errors.Add(string.Format("[ProcessItem] Validation issue with field {0}, <span style='color:#660000;'>{1}</span>", info.Name, info.Status));
                        }

                        if (result.Count == 0)
                        {
                            bool emailFailedToSend = false; // keep track, since there are quite a few different paths
                            string emailFailedReason = string.Empty; // friendly reason, which may be delivered to business folk
                            string fileOnDisk = string.Empty;
                            string exchangeMessage = string.Empty;

                            if (data.PartnerContactEmail.IsValidEmailAddress())
                            {
                                try
                                {
                                    w.Save();
                                    if (!ExistingMatchesAlreadyKnown.ContainsKey(data.VirtualKey))
                                    {
                                        ExistingMatchesAlreadyKnown.Add(data.VirtualKey, data.Id);
                                    }
                                    EmailGeneric(data, out fileOnDisk, displayLang);
                                    w.History += string.Format("[SYSTEM] An email was sent :) see attached for our record of the communication.'<br/>");
                                }
                                catch (Exception ex)
                                {
                                    emailFailedToSend = true;
                                    w.Fields["Microsoft.Operations.EmailControl.InvitationReminder"].Value = "Error";
                                    emailFailedReason = string.Format("When attempting message delivery, a system error was encountered: '<i>{0}</i>'.", ex.Message);
                                    w.History += string.Format("[SYSTEM] Tried to send an email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                                }
                            }
                            else
                            {
                                emailFailedToSend = true;
                                emailFailedReason = string.Format("The email address provided doesn't contain proper characters '<i>{0}</i>'.", data.PartnerContactEmail);
                                w.Fields["Microsoft.Operations.EmailControl.InvitationReminder"].Value = "Error";
                                w.History += Environment.NewLine + string.Format("[SYSTEM] Service Account wanted to send a confirmation email, but couldn't. The main email given, does not appear to be a properly formed address '<i>{0}</i>'.", data.PartnerContactEmail);
                            }

                            if (!string.IsNullOrEmpty(fileOnDisk))
                            {
                                try
                                {
                                    // Also add, as an attachment, original spreadsheet submission.
                                    Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileOnDisk, "Confirmation Email");
                                    w.Attachments.Add(eml);
                                }
                                catch (Exception ex)
                                {
                                    //Log.WriteLine(string.Format("ERROR: Tried to add attachment '{0}' to SKU Purchase '{1}', but failed. More info: {2}", fileOnDisk, data.VirtualKey, ex.StackTrace));
                                }
                            }

                            if (!string.IsNullOrEmpty(data.OriginalFragment))
                            {
                                // For audit trail purposes, we would like to also attach a copy of
                                // the xml fragment. This will help any analysis in the event there
                                // is an estimated translation error. Required because there is no
                                // field-data validation on the actual RegSys form - i.e. data
                                // entered can be anything. Since it's gone live, we've already seen
                                // some weird values come through.

                                // TODO: This file-creation-from-string-plus-attachment should be
                                //       made into a COMMON ROUTINE for simplicity Ensure the
                                // currently executing thread has permissions to create the folder if
                                // it does not exist already.

                                string fileFragmentFolder = string.Format(@"{0}\RegSys\Temp", FileSystem.BaseFolder);
                                string fileFragmentName = string.Format(@"{0}\RegSysRawData.{1}.{2:yyyyMMdd.HHmmss.fff}.txt", fileFragmentFolder, data.VirtualKey, DateTime.Now);

                                if (!Directory.Exists(fileFragmentFolder))
                                {
                                    Directory.CreateDirectory(fileFragmentFolder);
                                }

                                StreamWriter temp = new StreamWriter(fileFragmentName);
                                temp.WriteLine(data.OriginalFragment);
                                temp.Flush();
                                temp.Close();
                                temp.Dispose();

                                Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileFragmentName, "RegSys Submission RAW Data");
                                w.Attachments.Add(eml);
                            }
                            try
                            {
                                w.Save();
                            }
                            catch (Exception ex)
                            {
                                //Log.WriteLine(string.Format("Failed to send an administrative email warning of email error to send - associated. The item in question was {0}, and the actual error was '{1}' ... {2}", data.VirtualKey, ex.Message, ex.Source));
                            }
                            // Escalate by letting an administrator know. Note that we only do this here
                            if (emailFailedToSend)
                            {
                                try
                                {
                                    StringBuilder content = new StringBuilder();
                                    content.AppendLine("An error occurred with sending an email from the RegSys information<br/>");
                                    content.AppendLine("The error was: '<b>{0}</b>'<br/>", emailFailedReason);
                                    content.AppendLine("And the associated record in CSP is #{0}", w.Id);
                                    content.AppendLine("Please note the confirmation email may need to be generated and sent manually.");

                                    string bodyText = FileSystem.GetEmbeddedFileContent("Microsoft.Operations.CSP.RegSys.Templates.Blank.html", "Microsoft.Operations.CSP.RegSys");
                                    string compressionResult = string.Empty;

                                    bodyText = Optimize.MinifyHtml(bodyText, out compressionResult, true);

                                    bodyText = bodyText.Replace("###MAIN_CONTENT###", content.ToString());

                                    ExchangeService svc = ExchangeServiceAutomatic.New(EmailUserIdentityEmail, EmailPassword);
                                    EmailMessage message = new EmailMessage(svc);

                                    message.Subject = string.Format("Microsoft Advanced Support for Partners (ASfP) Email Delivery Failure");
                                    message.ToRecipients.Add("svcasfp@microsoft.com");
                                    //message.BccRecipients.Add("cboinfrastructure@microsoft.com");
                                    //message.BccRecipients.Add("chads@microsoft.com");
                                    message.ReplyTo.Add("svcasfp@microsoft.com");
                                    message.Body = new MessageBody(BodyType.HTML, bodyText);
                                    Email.InsertImageFromResource(ref message, "microsoft_footer.png");

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
                                            Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment eml = new TeamFoundation.WorkItemTracking.Client.Attachment(fileNamePayload, "Error Email");
                                            w.Attachments.Add(eml);
                                        }
                                        catch (Exception ex)
                                        {
                                            w.History += string.Format("[SYSTEM] Tried to attach email, but failed! The error given by Exchange was: '<i>{0}</i>'", ex.Message);
                                        }
                                    }
                                    w.Save();
                                }
                                catch (Exception ex)
                                {
                                    //Log.WriteLine(string.Format("Failed to send an administrative email warning of email error to send - associated. The item in question was {0}, and the actual error was '{1}' ... {2}", data.VirtualKey, ex.Message, ex.Source));
                                }
                            }
                            try
                            {
                                w.Save();
                            }
                            catch (Exception ex)
                            {
                                w.History += string.Format("[SYSTEM] Failed to save work item: '<i>{0}</i>'", ex.Message);
                            }
                        }
                    } // end of testing if this RegSys item is already known to have a match
                }
            } // end of looping through identified entries from RegSys
        }

        public void ProcessWizardEntries_InvokeThread()
        {
            if (ProcessWizardEntries != null && !ProcessWizardEntries.IsBusy)
            {
                ProcessWizardEntries.RunWorkerAsync();
            }
        }

        // end of doing work/thread
    }
}