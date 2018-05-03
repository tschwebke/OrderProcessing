using Microsoft.WindowsAzure.Storage.Table;
using System;

/// <summary>
/// Same as an event log item, but in a cloud table.
/// *** THIS IS THE OLD 2014-2015 version WHICH WILL BE PHASED OUT OVER TIME (use the 2016 version
/// instead: 'LogDetail')
/// </summary>
public class LogEntry : TableEntity
{
    /// <summary>
    /// This parameterless constructor needs to stay - requirement of cloud object.
    /// </summary>
    public LogEntry() { }

    public LogEntry(string partitionKey)
    {
        PartitionKey = partitionKey;
        RowKey = string.Format("{0:yyyyMMdd.HHmmss.fff}", DateTime.UtcNow);
    }

    /// <summary>
    /// Main text of the event/message, etc.
    /// </summary>
    public string EventData { get; set; }

    /// <summary>
    /// Information, Warning, Error or Critical
    /// </summary>
    public string Level { get; set; }

    /// <summary>
    /// From what hosting machine, is the entry from?
    /// </summary>
    public string MachineName { get; set; }

    /// <summary>
    /// Name of the executing identity (i.e. user name) of the MAIN thread of the running program.
    /// </summary>
    public string SecurityContext { get; set; }
}