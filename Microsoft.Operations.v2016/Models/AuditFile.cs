using Microsoft.WindowsAzure.Storage.Table;
using System;

/// <summary>
/// Basic detail of an Email which was sent or received by the system.
/// </summary>
public class AuditFile : TableEntity
{
    /// <summary>
    /// This parameterless constructor needs to stay - requirement of cloud object.
    /// </summary>
    public AuditFile() { }

    public AuditFile(string partitionKey)
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
    /// What is the nature of this audit entry? i.e. to help reporting and filtering
    /// - 'AttachmentReceipt'
    /// </summary>
    public string AuditType { get; set; }

    /// <summary>
    /// Email address (SMTP) of who the email object was from. This is NOT related to the internal
    /// content of the email.
    /// </summary>
    public string From { get; set; }

    /// <summary>
    /// Record of the primary internet headers from the original message, the preservation of which
    /// may aid troubleshooting and security diagnosis. [ This is LOW VALUE information and we may
    /// consider dropping this data in future versions ]
    /// </summary>
    public string Headers { get; set; }

    /// <summary>
    /// The alias-based email address denoting the mailbox which handled this email.
    /// </summary>
    public string Mailbox { get; set; }
}