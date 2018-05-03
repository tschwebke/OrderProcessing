using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.WindowsAzure.Storage;
using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage.Table;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        private Dictionary<string, int> ExistingMatchesAlreadyKnown = new Dictionary<string, int>();

        public async void ProcessWizardEntries_DoWork(object sender, DoWorkEventArgs e)
        {
            Collection<RegSysWizardItem> recent_submissions = new Collection<RegSysWizardItem>();

            DateTime today = DateTime.Today.AddDays(1); // go a day ahead to avoid any international time conflicts

            // Test Url for Azure Table Storage Resr API
            //https://asfptbl.table.core.windows.net/asfptable()?sv=2017-07-29&ss=bfqt&srt=sco&sp=rwdlacup&se=2018-04-01T04:31:28Z&st=2018-03-07T21:31:28Z&spr=https&sig=6lqJgIksvEMqe%2BVGPB5ox%2BWmmx6MnVRqzA2xPTKQKf8%3D

            string Url = "https://asfptbl.table.core.windows.net/asfptable";

            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(Url);
            webRequest.Method = "GET";
            webRequest.ContentLength = 0;
            webRequest.Headers.Add("x-ms-date", DateTime.UtcNow.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
            string resource = "asfptable";
            string account = "asfptbl";
            string secret = "h4bDfcQ/NJAH2CiyN2ruqBBfbQgb2/1LHVsT8rpDJ0v2ybDEiWLw/wWNxyKkZUVRc6S5Z/VJGObDV5hQLOnOsA==";

            //TODO: Limit the number of records returned

            // Now sign the request For a table, you need to use this Canonical form: VERB + "\n" +
            // Content - MD5 + "\n" + Content - Type + "\n" + Date + "\n" + CanonicalizedResources; Verb
            string signature = "GET\n";

            // Content-MD5
            signature += "\n";

            // Content-Type
            signature += "\n";

            // Date
            signature += webRequest.Headers["x-ms-date"] + "\n";

            // Canonicalized Resource remove the query string
            int q = resource.IndexOf("?");
            if (q > 0) resource = resource.Substring(0, q);

            // Format is /{0}/{1} where 0 is name of the account and 1 is resources URI path
            signature += "/" + account + "/" + resource;

            // Hash-based Message Authentication Code (HMAC) using SHA256 hash
            System.Security.Cryptography.HMACSHA256 hasher = new System.Security.Cryptography.HMACSHA256(Convert.FromBase64String(secret));

            // Authorization header
            string authH = "SharedKey " + account + ":" + System.Convert.ToBase64String(hasher.ComputeHash(System.Text.Encoding.UTF8.GetBytes(signature)));

            // Add the Authorization header to the request
            webRequest.Headers.Add("Authorization", authH);
            string html;

            HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse();
            using (System.IO.StreamReader r = new System.IO.StreamReader(response.GetResponseStream()))
            {
                html = r.ReadToEnd();
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(html);

            var mgr = new XmlNamespaceManager(xmlDoc.NameTable);
            mgr.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
            mgr.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");
            mgr.AddNamespace("georss", "http://www.georss.org/georss");
            mgr.AddNamespace("gml", "http://www.opengis.net/gml");

            // The conversion of the raw data to a RegSysWizardItem has some unique and subtle
            // considerations to it. If you are working extensively with the object, be sure to view
            // the notes (inside 'RegSysWizardItem.cs')

            foreach (XmlElement x in xmlDoc.SelectNodes("//m:properties", mgr))
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
                    if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "thads")
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
                        w.Fields["Microsoft.Operations.Partners.Support.GeographicalPreference"].Value = data.PartnerArea.Trim();
                        w.Fields["Microsoft.Operations.Partners.Contact.Name"].Value = data.PartnerContactName.Trim();
                        w.Fields["Microsoft.Operations.Partners.Contact.Email"].Value = data.PartnerContactEmail.Trim();
                        w.Fields["Microsoft.Operations.Partners.Partner.PhoneNumber"].Value = data.PartnerPhone.Trim();
                        w.Fields["Microsoft.Operations.CreditDiscountApproval.Request.SKUs"].Value = data.SKUs;
                        w.Fields["Microsoft.Operations.Partners.Partner.CurrentState"].Value = @"ASfP Purchase";
                        //w.Fields["Microsoft.Operations.Partners.Billing.ContactName"].Value = data.BillingContactName.Trim();
                        //w.Fields["Microsoft.Operations.Partners.Billing.ContactEmail"].Value = data.BillingContactEmail.Trim();
                        w.Fields["Microsoft.Operations.Partners.AreaOfExpertise"].Value = data.PartnerArea.Trim();
                        w.Fields["Microsoft.Operations.AutomatedProcess.LastChecked"].Value = data.CreatedDateInRegSys; // store here for now ...
                        w.Fields["Microsoft.Operations.Partners.ExternalIdentifiers.MPNID"].Value = data.MPNID.Trim();
                        w.Fields["Microsoft.Operations.Partners.Nomination.OriginalSource"].Value = data.DataOrigin;
                        w.Fields["System.AssignedTo"].Value = @""; // default

                        string Region;
                        switch (data.PartnerArea.Trim())
                        {
                            case "アジア太平洋":
                                Region = "Asia Pacific";
                                break;

                            case "ヨーロッパ、中東、アフリカ":
                                Region = "Europe, Middle East, Africa";
                                break;

                            case "南米":
                                Region = "Latin America";
                                break;

                            case "日本":
                                Region = "Japan";
                                break;

                            case "北米":
                                Region = "North America";
                                break;

                            default:
                                Region = data.PartnerArea.Trim();
                                break;
                        }

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
                        w.Fields["Microsoft.Operations.Partners.FormLanguage"].Value = displayLanguage;

                        if (!string.IsNullOrEmpty(data.CustomerTenantDomain))
                        {
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Company.TenantDomain"].Value = data.CustomerTenantDomain.Trim();
                        }
                        else
                        {
                            w.Fields["Microsoft.Operations.CreditDiscountApproval.Company.TenantDomain"].Value = "";
                        }

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
                                    // Retrieve the storage account from the connection string.
                                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
                                        CloudConfigurationManager.GetSetting("asfptbl_AzureStorageConnectionString"));

                                    // Create the table client.
                                    CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

                                    // Create the CloudTable that represents the "people" table.
                                    CloudTable table = tableClient.GetTableReference("asfpTable");

                                    // Create a retrieve operation that expects a customer entity.
                                    TableOperation retrieveOperation = TableOperation.Retrieve<DynamicTableEntity>(Region, data.VirtualKey);

                                    // Execute the operation.
                                    TableResult retrievedResult = table.Execute(retrieveOperation);

                                    // Assign the result to a CustomerEntity.
                                    DynamicTableEntity deleteEntity = (DynamicTableEntity)retrievedResult.Result;

                                    // Create the Delete TableOperation.
                                    if (deleteEntity != null)
                                    {
                                        TableOperation deleteOperation = TableOperation.Delete(deleteEntity);
                                        table.Execute(deleteOperation);

                                        Console.WriteLine("Entity deleted.");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Could not retrieve the entity.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //Log.WriteLine(string.Format("Failed to send an administrative email warning of email error to send - associated. The item in question was {0}, and the actual error was '{1}' ... {2}", data.VirtualKey, ex.Message, ex.Source));
                                }
                            }
                            try
                            {
                                w.Save();
                                // Retrieve the storage account from the connection string.
                                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
                                    CloudConfigurationManager.GetSetting("asfptbl_AzureStorageConnectionString"));

                                // Create the table client.
                                CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

                                // Create the CloudTable that represents the "people" table.
                                CloudTable table = tableClient.GetTableReference("asfpTable");

                                // Create a retrieve operation that expects a customer entity.
                                TableOperation retrieveOperation = TableOperation.Retrieve<DynamicTableEntity>(Region, data.VirtualKey);

                                // Execute the operation.
                                TableResult retrievedResult = table.Execute(retrieveOperation);

                                // Assign the result to a CustomerEntity.
                                DynamicTableEntity deleteEntity = (DynamicTableEntity)retrievedResult.Result;

                                // Create the Delete TableOperation.
                                if (deleteEntity != null)
                                {
                                    TableOperation deleteOperation = TableOperation.Delete(deleteEntity);
                                    table.Execute(deleteOperation);

                                    Console.WriteLine("Entity deleted.");
                                }
                                else
                                {
                                    Console.WriteLine("Could not retrieve the entity.");
                                }
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
    }
}