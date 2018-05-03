using Microsoft.WindowsAzure.Storage.Table;
using System;

/// <summary>
/// Basic detail of an Email which was sent or received by the system. This object is one in a series
/// of 'Audit' items which are permanent records of semi-important actions. They are all stored
/// inside the cboservicelogs 'audit' table.
/// </summary>
public class AuditEmail : TableEntity
{
    /// <summary>
    /// This parameterless constructor needs to stay - requirement of cloud object.
    /// </summary>
    public AuditEmail() { }

    public AuditEmail(string partitionKey)
    {
        this.PartitionKey = partitionKey;
        this.RowKey = string.Format("{0:yyyyMMdd.HHmmss.fff}", DateTime.UtcNow);
    }

    /// <summary>
    /// The action applied TO the email object AFTER it was used. i.e. this gives an indication of
    /// where the actual data might be kept, but is agnostic of any other activity.
    /// - 'MovedToSpecificFolder'
    /// - 'RetainedCopyInSentItems'
    /// - 'MovedToRecycleBin'
    /// - 'HardDeleted'
    /// - 'NoCopyKept'
    /// </summary>
    public string ArchiveAction { get; set; }

    /// <summary>
    /// What is the nature of this audit entry? What event is it describing? i.e. to help reporting
    /// and filtering
    /// - 'EmailDelivery' : For record of outgoing emails, where we need a record
    /// - 'EmailReceived' : For recording incoming items (and subsequent processing)
    /// </summary>
    public string AuditType { get; set; }

    /// <summary>
    /// Email address (SMTP) of who the email object was from. This is NOT necessarily related to the
    /// internal content of the email.
    /// </summary>
    public string From { get; set; }

    /// <summary>
    /// Record of the primary internet headers from the original message, the preservation of which
    /// may aid troubleshooting and security diagnosis. [ This is LOW VALUE information and we may
    /// consider dropping this data in future versions ]
    /// </summary>
    public string InternetHeaders { get; set; }

    /// <summary>
    /// From what hosting machine, is the entry from?
    /// </summary>
    public string MachineName { get; set; }

    /// <summary>
    /// The alias-based email address denoting the mailbox which handled this email.
    /// </summary>
    public string Mailbox { get; set; }

    /// <summary>
    /// Name of the executing identity (i.e. user name) of the MAIN thread of the running program.
    /// </summary>
    public string SecurityContext { get; set; }

    /// <summary>
    /// The title on the original email object (may also be left blank). This is to help
    /// identify/match any archived objects in the mailbox, if required.
    /// </summary>
    public string Subject { get; set; }

    /// <summary>
    /// Free text - the name of the function which handled this particular item. Use the method (or
    /// class) name from code as a handy default if a specific namer is not available. If applicable,
    /// separate multiple stages with full stops so that grouping is clear, e.g. for the first part
    /// of processing, "WorkflowX.Part1" etc.
    /// </summary>
    public string WorkflowStage { get; set; }
}