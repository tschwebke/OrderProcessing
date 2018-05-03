using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        // IMPORTANT NOTE: THIS IS SAMPLE CODE ONLY, FOR READING AN INBOX

        public void ProcessInbox_InvokeThread()
        {
            //if (ProcessInbox != null && !ProcessInbox.IsBusy)
            //{
            //    ProcessInbox.RunWorkerAsync();
            //}
        }

        /// <summary>
        /// Check the inbox for any spreadsheet-based submissions.
        /// </summary>
        public void ProcessInbox_DoWork(object sender, DoWorkEventArgs e)
        {
            //AzureLogger Log = new AzureLogger();

            ExchangeService svc = ExchangeServiceAutomatic.New(UserIdentityEmail, ServicePassword);

            PropertySet itempropertysetBasic = new PropertySet(BasePropertySet.FirstClassProperties);
            PropertySet itempropertysetExtended = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, ItemSchema.Attachments) { RequestedBodyType = BodyType.Text };  // gets all various properties, not just the shell ... http://msdn.microsoft.com/en-us/library/aa566107%28v=exchg.140%29.aspx

            ItemView itemview = new ItemView(1000);
            itemview.PropertySet = itempropertysetBasic;

            FindItemsResults<Microsoft.Exchange.WebServices.Data.Item> findResults = svc.FindItems(WellKnownFolderName.Inbox, itemview);

            // This code gets information about the folder structure, if required. Folder rootfolder
            // = Folder.Bind(svc, WellKnownFolderName.Inbox); foreach (Folder folder in
            // rootfolder.FindFolders(new FolderView(100))) { Console.WriteLine("\nName: " +
            // folder.DisplayName + "\n Id: " + folder.Id); }

            // These folders remain in the inbox as permanent records of what was sent. This is handy
            // because we get the ability to replay later if there are any data issues.

            string web_leads_folder_id = "AAMkADRkZDVjMmY4LTgyYmMtNDYzMy04MTg5LTM1NjAyZWE3MjQyYgAuAAAAAACrtq4Z1bVgRo3DFLNydbMHAQDe16qv4GtLSJbtRVUIX8iSAAAE6ex0AAA=";
            string spreadsheet_folder_id = "AAMkADRkZDVjMmY4LTgyYmMtNDYzMy04MTg5LTM1NjAyZWE3MjQyYgAuAAAAAACrtq4Z1bVgRo3DFLNydbMHAQDe16qv4GtLSJbtRVUIX8iSAAAE6ex3AAA=";

            foreach (Microsoft.Exchange.WebServices.Data.Item item in findResults.Items)
            {
                item.Load(itempropertysetExtended);

                string subjectLine = Optimize.SafeString(item.Subject);

                EmailMessage email = (EmailMessage)item;
                EmailAddress senderEmail = email.From; // can only access this value if ALL properties have been loaded.

                // TODO: Use some case statements for this, as well as a common handler.

                string junkMailReason = string.Empty;
                int junkConfidenceLevel = 0;

                if (email.IsJunkMail(out junkMailReason, out junkConfidenceLevel))
                {
                    // We will treat the junk mail item on the basis of the confidence of the decision.

                    if (junkConfidenceLevel >= 100)
                    {
                        //Log.WriteLine(string.Format("removing junk email from [{0}] '{1}' HARD DECISION: {2}", senderEmail.Address, email.Subject, junkMailReason));
                        item.Delete(DeleteMode.HardDelete);
                    }
                    else
                    {
                        //Log.WriteLine(string.Format("moving possible junk email to 'deleted items' from [{0}] '{1}' SOFT DECISION: {2}", senderEmail.Address, email.Subject, junkMailReason));
                        item.Delete(DeleteMode.MoveToDeletedItems);
                    }
                }
                else if (subjectLine.StartsWith("Undeliverable: "))
                {
                    // Get the date of this item ... extract the email address from the body text ...

                    string sample = item.Body.Text.Substring(item.Body.Text.IndexOf("mailto:") + 7, 250); // should contain text: "Delivery has failed to these recipients or groups"
                    string failedEmail = sample.Substring(0, sample.IndexOf(@">") - 1);

                    // TODO: Write to cloud table for later handling on 'failed deliveries'

                    item.Delete(DeleteMode.HardDelete);
                }
                else if (subjectLine == "Partner Center CSP Inquiry" && senderEmail.Name == "Cloud Solution Provider (CSP) Program")
                {
                    try
                    {
                        email.IsRead = true;
                        email.Update(ConflictResolutionMode.AlwaysOverwrite);
                        email.Move(new FolderId(web_leads_folder_id));
                    }
                    catch (Exception ex)
                    {
                        email.Delete(DeleteMode.MoveToDeletedItems);
                    }
                }
                else if (item.HasAttachments)
                {
                    // Note: to distinguish between types of attachments (attached v.s. inline
                    // images) use the following CID method: http://stackoverflow.com/questions/6650842/ews-exchange-2007-retrieve-inline-images

                    foreach (var i in item.Attachments)
                    {
                        FileAttachment fileAttachment = i as FileAttachment;

                        if (fileAttachment.Name.ToLower().Contains(".xls")) // picks up both types of Excel file (identifies both legacy & modern)
                        {
                            // As of October 2015, there are more than one streams depending on the
                            // nature of attachment received.

                            if (fileAttachment.Name.ToLower().Contains("discount") || fileAttachment.Name.ToLower().Contains("google") || fileAttachment.Name.ToLower().Contains("promo"))
                            {
                                // Execute 'discount' flow. REF: Task 63957

                                string saveFileAs = Path.Combine(Path.GetTempPath(), fileAttachment.Name);

                                fileAttachment.Load(saveFileAs); // Save to disk in temp folder

                                FileInfo spreadsheetOriginal = new FileInfo(saveFileAs); // TODO: Make generic to a 'FileProcessing' object
                                Collection<ErrorDetail> itemsRequiringAttention = new Collection<ErrorDetail>(); // TODO: Make generic to a 'FileProcessing' object

                                string bodyContent = Optimize.SafeString(email.Body).Replace("\r\n", "<br/>");

                                // Discount newSubmissionViaSpreadsheet =
                                // ProcessDiscountSubmission(Log, spreadsheetOriginal, senderEmail,
                                // subjectLine, bodyContent);

                                // EmailSpreadsheetProcessingResultDiscount(spreadsheetOriginal,
                                // senderEmail, newSubmissionViaSpreadsheet);

                                spreadsheetOriginal.Delete(); // The spreadsheet may contain partner information, we absolutely don't want the original record floating around on the server file system.

                                foreach (ErrorDetail ed in itemsRequiringAttention)
                                {
                                    //Log.WriteLine(string.Format("Error: {0}", ed.FriendlyMessage));
                                }
                            }
                            else
                            {
                                // NORMAL spreadshet for CSP Nominations
                                // TODO: Add this to a separate routine/handler

                                // First, need to go and get a list of users who are qualified to
                                // send a spreadsheet (see Task 42244). PSE Account Managers are also
                                // allowed to submit. Use this design pattern for any additional
                                // groups (ref. Task 50591).

                                // Dictionary<string, string> authorizedUsers =
                                // TfsHelper.LookupGroupMembership(ServerName, "PBD Leads");
                                // Dictionary<string, string> authorizedUsers_additional =
                                // TfsHelper.LookupGroupMembership(ServerName, "PSE Account Managers");

                                Dictionary<string, string> authorizedUsers = new Dictionary<string, string>();
                                Dictionary<string, string> authorizedUsers_additional = new Dictionary<string, string>();

                                authorizedUsers.MergeWith(authorizedUsers_additional);

                                // Note the above structure obtains only the alias@microsoft.com
                                // format, it doesn't have knowledge of the easyid format. ref: Task
                                // 49031: 'EasyID' support.

                                if (!authorizedUsers.ContainsValue(senderEmail.Address.ToLower()))
                                {
                                    // The user does not have the permission to send a spreadsheet .. EmailWrongPermissionsReponse(senderEmail);
                                    item.Delete(DeleteMode.MoveToDeletedItems);
                                }
                                else
                                {
                                    // process normal

                                    string saveFileAs = Path.Combine(Path.GetTempPath(), fileAttachment.Name);

                                    //Log.WriteLine(string.Format("Exchange Inbox - location: {0} found {1} items", svc.Url, findResults.TotalCount));
                                    //Log.WriteLine(string.Format("An attachment '{1}' has been sent to the service account! (processing it @ {2}).", Environment.NewLine, fileAttachment.Name, saveFileAs));

                                    fileAttachment.Load(saveFileAs); // Save to disk in temp folder

                                    FileInfo spreadsheetOriginal = new FileInfo(saveFileAs); // reference to the file just saved
                                    Collection<ErrorDetail> itemsRequiringAttention = new Collection<ErrorDetail>(); // collection for any potential errors

                                    string bodyContent = Optimize.SafeString(email.Body).Replace("\r\n", "<br/>");

                                    // RegSysWizardItem newSubmissionViaSpreadsheet =
                                    // ProcessSpreadsheetSubmission(Log, spreadsheetOriginal,
                                    // senderEmail, subjectLine, bodyContent); // for storing
                                    // grid-based contents of original submission

                                    // Finally, give feedback to the original sender about what they
                                    // sent (good or bad).

                                    // Non-compliant No permissions

                                    // EmailSpreadsheetProcessingResult(spreadsheetOriginal,
                                    // senderEmail, newSubmissionViaSpreadsheet);

                                    try
                                    {
                                        // ArchiveToVersionControl(spreadsheetOriginal,
                                        // string.Format("spreadsheet submission received from:
                                        // '{0}'", senderEmail), ref itemsRequiringAttention, ref LogInbox);
                                    }
                                    catch (Exception ex)
                                    {
                                        //Log.WriteLine(string.Format("Error Saving to version control Email: {0}", ex.Message));
                                    }
                                    finally
                                    {
                                        spreadsheetOriginal.Delete(); // The spreadsheet may contain partner information, we absolutely don't want the original record floating around on the server file system.
                                    }

                                    foreach (ErrorDetail ed in itemsRequiringAttention)
                                    {
                                        //Log.WriteLine(string.Format("Error: {0}", ed.FriendlyMessage));
                                    }
                                } // end of authorization testing.
                            } // end of testing spreadsheet name
                        } // end of testing if attachment is excel.
                    } // if the email has attachments.

                    // Handles multiple attachments. i.e. don't move the email object until we have
                    // processed it !!!

                    try
                    {
                        // item.Load(New PropertySet({ EmailMessageSchema.IsRead}))
                        email.IsRead = true;
                        email.Update(ConflictResolutionMode.AlwaysOverwrite);
                        email.Move(new FolderId(spreadsheet_folder_id));
                    }
                    catch
                    {
                        email.Delete(DeleteMode.MoveToDeletedItems);
                    }
                }
                else
                {
                    // for anything else ... e.g. random email ... we have to tell them about what
                    // this alias is used for.
                }
            }
        }
    }
}