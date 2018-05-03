using Microsoft.Exchange.WebServices.Data;
using Microsoft.Operations;
using System;
using System.IO;
using System.Linq;

/// <summary>
/// Container which can be used for common email delivery. Centralized error catching, logging and
/// file-system transactions.
/// REQUIRES: File System Access to temp directory.
/// </summary>
public class EmailMessageLauncher : IDisposable
{
    /// <summary>
    /// Comments which would accompany the file attachment, e.g. if it's used to save into a TFS work
    /// item. Default is set, but can be overridden if necessary.
    /// </summary>
    public string AttachmentComments;

    public EmailMessage Email;

    public int ErrorCount;

    public string ErrorMessage;

    /// <summary>
    /// Name of a temp file which gets created in local 'temp' folder. Automatic deletion if used
    /// with the 'attach' method.
    /// </summary>
    public string FileNameEmailPayload;

    /// <summary>
    /// Azure-based logging!
    /// </summary>
    public AzureLogger Logger;

    private string FileNamePatternTemplate;

    /// <summary>
    /// </summary>
    /// <param name="log"></param>
    /// <param name="msg">
    /// Specifies an email to use (it should already be fully formed prior to loading)
    /// </param>
    /// <param name="fileNamePattern">A base name to use for the final file name e.g. 'email_tier2_invite.eml'</param>
    public EmailMessageLauncher(AzureLogger log, EmailMessage msg, string fileNamePattern)
    {
        ErrorCount = 0;
        Email = msg;
        Logger = log;
        AttachmentComments = string.Format("Email Communication {0:dd-MMM-yyyy}", DateTime.Now);
        FileNamePatternTemplate = fileNamePattern;
        FileNameEmailPayload = Path.Combine(Path.GetTempPath(), FileNamePatternTemplate.Replace(".eml", string.Format("_{0}.eml", string.Join("", Guid.NewGuid().ToString("n").Take(4).Select(o => o))))); // target is the native temp directory (always available) ... appends random chars for uniqueness.
    }

    /// <summary>
    /// CLEANUP! Note: if using this with a TFS workitem, only do this after the work item has been
    /// saved, otherwise an error will occur.
    /// </summary>
    public void Dispose()
    {
        // if the temp file still exists at the point, clean it out.
        if (File.Exists(FileNameEmailPayload))
        {
            File.Delete(FileNameEmailPayload);
        }
    }

    /// <summary>
    /// Basic validation on the email to see if it can be sent. Will set a value for 'ErrorMessage'
    /// with a basic description.
    /// </summary>
    public bool IsReadyToSend()
    {
        bool isGood = true;
        ErrorMessage = string.Empty;

        if (Email.ToRecipients.Count == 0)
        {
            ErrorMessage = "No receipients specified!";
            ErrorCount++;
            isGood = false;
        }

        if (string.IsNullOrEmpty(Email.Subject))
        {
            ErrorMessage = "Missing subject line!";
            ErrorCount++;
            isGood = false;
        }

        foreach (EmailAddress a in Email.ToRecipients)
        {
            if (!a.Address.IsValidEmailAddress())
            {
                ErrorMessage += " Malformed email address: " + a.Address;
                ErrorCount++;
                isGood = false;
            }
        }

        return isGood;
    }

    /// <summary>
    /// Saves a file
    /// </summary>
    public void SendWithAuditTrail(bool createLocalFileCopyForLaterAttachment = true, bool saveInSentItemsFolder = true)
    {
        if (createLocalFileCopyForLaterAttachment)
        {
            Email.Save(WellKnownFolderName.Drafts); // this is required to get the "ID" value so we can access other properties of the object. After the mail gets sent, this is automatically removed from DRAFTS

            Email.Load(new PropertySet(ItemSchema.MimeContent));
            var mimeContent = Email.MimeContent;

            using (var fileStream = new FileStream(FileNameEmailPayload, FileMode.Create))
            {
                fileStream.Write(mimeContent.Content, 0, mimeContent.Content.Length);
            }
        }

        // TODO: What happens if an error occurs here, with the actual sending?

        if (saveInSentItemsFolder)
        {
            Email.SendAndSaveCopy(WellKnownFolderName.SentItems); // keeps a copy here
        }
        else
        {
            Email.Send(); // normal send - but no record is kept (suitable for high-volume emails)
        }
    }
}