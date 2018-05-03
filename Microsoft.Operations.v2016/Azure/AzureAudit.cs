using Microsoft.Exchange.WebServices.Data;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using System;

/// <summary>
/// One in a series of 'Audit' items which are permanent records of semi-important actions.
/// </summary>
public class AzureAudit
{
    private string _assembly_name = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
    private CloudTable _audit_store;
    private string _executing_identity = System.Environment.UserName.ToLower();
    private string _machine_name = System.Environment.MachineName.ToUpper();
    private CloudStorageAccount _storage_account;

    /// <summary>
    /// An 'Audit' object is extended information about an action or data that allows us to trace
    /// important activity. These entries get retained indefinitely (retention subject to regular
    /// review). Similar to logging but more 'permanent'. These objects will be used for reporting.
    /// </summary>
    public AzureAudit()
    {
        _storage_account = CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=cboservicelogs;AccountKey=7psl9DFh8F3AISj1FoqwtMyBNQm4ni7r2ES1tF5PiEsJgxDbYzZ7LFDCLQHBU0Hmv2NN+we2WmJfbTP12VxClA==");
        _audit_store = _storage_account.CreateCloudTableClient().GetTableReference("audit");
    }

    /// <summary>
    /// Full 'write entry' which accepts arguments.
    /// NOTE: Be sure to have loaded ALL THE PROPERTIES of the email object before they get accessed,
    ///       otherwise a null reference exception may occur.
    /// </summary>
    public bool TimestampIncomingCommunication(EmailMessage email, string mailboxName, string methodName, string archiveAction)
    {
        bool operationWasSuccessful;
        AuditEmail stamp = new AuditEmail(_assembly_name);

        stamp.AuditType = "EmailReceived"; // never changes, for incoming emails being processed.
        stamp.From = email.From.Address;
        stamp.Subject = email.Subject;
        stamp.WorkflowStage = methodName; // e.g. "ProcessWebLeadEmail";
        stamp.MachineName = _machine_name;
        stamp.Mailbox = mailboxName;
        stamp.SecurityContext = _executing_identity;
        stamp.ArchiveAction = archiveAction; // e.g. "MovedToSpecificFolder";

        foreach (InternetMessageHeader h in email.InternetMessageHeaders)
        {
            if (!h.Name.StartsWith("X-")) stamp.InternetHeaders += h.Name + ": " + h.Value + Environment.NewLine; // Ignore any extended properties, not required.
        }

        try
        {
            TableOperation update = TableOperation.InsertOrReplace(stamp);
            _audit_store.Execute(update);
            operationWasSuccessful = true;
        }
        catch (Exception ex)
        {
            operationWasSuccessful = false;
        }

        return operationWasSuccessful;
    }
}