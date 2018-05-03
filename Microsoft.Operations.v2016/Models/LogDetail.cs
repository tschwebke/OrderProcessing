using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Linq;

/// <summary>
/// 2016 version - please use this for logging all application activity. PartitionKey = "number of
/// days until x" bucket, for day-based logging (and for easier deletion) RowKey = Logtail pattern
/// using Ticks, which enables easy TOP-x based queries (ref: https://azure.microsoft.com/en-us/documentation/articles/storage-table-design-guide/#table-design-patterns)
/// </summary>
public class LogDetail : TableEntity
{
    /// <summary>
    /// This parameterless constructor needs to stay - requirement of cloud object.
    /// </summary>
    public LogDetail() { }

    /// <summary>
    /// Create a new log entry to jam free-form information. Similar to local EventLog, but better
    /// because it's cloud-based. Stored in 'ApplicationChatter' table, suitable for recording any event.
    /// </summary>
    /// <param name="applicationName">
    /// Typically the namespace of the service e.g. Microsoft.Operations.WHT.Service, but may also be
    /// abbreviated form.
    /// </param>
    /// <param name="originalCallingMethodName">
    /// Optional, describes where the logger where in the program this is occuring (for identification)
    /// </param>
    /// <param name="codeFile">
    /// Name of the (compile file) where the logging item gets called, for easy later troubleshooting.
    /// </param>
    /// <param name="lineNumber">(if known) the line number of the original caller</param>
    public LogDetail(string applicationName, string originalCallingMethodName, string codeFile, int lineNumber)
    {
        DateTime someFutureDate = new DateTime(2070, 1, 1); // date far in future to begin reverse ticks; this logging library should be long obsolete by that date

        Year = DateTime.UtcNow.Year;
        Month = DateTime.UtcNow.Month;
        Day = DateTime.UtcNow.Day;
        PartitionKey = someFutureDate.Subtract(DateTime.UtcNow).Days.ToString();
        RowKey = string.Format("{0:D17}{1}", someFutureDate.Ticks - DateTime.UtcNow.Ticks, string.Join("", Guid.NewGuid().ToString("n").Take(4).Select(o => o))); // logtail pattern with additional randomness
        CallerMemberName = originalCallingMethodName;
        CallerFilePath = codeFile;
        CallerLineNumber = lineNumber;
        ApplicationName = applicationName;

        // Notes regarding Azure Table Indexing choice as per above: PARTITIONKEY

        // ROWKEY There are already 10,000 ticks per millisecond so a collision of Rowkey (on the
        // same day) is ultimately unlikely, even with multiple services on the same machine using
        // the construct. But to make the item TRULY UNIQUE (and without affecting the order-by
        // value), we're appending random characters from a GUID. Only 4 values are selected but that
        // gives a possibility of 36^4 = 1.6 million random combinations within the same tick
        // instant, which is more than enough (maybe less depending on the method the CPU uses for
        // GUID generation, but near enough). There are also other ways of guaranteeing Tick
        // uniqueness, e.g.
        // http://stackoverflow.com/questions/5608980/how-to-ensure-a-timestamp-is-always-unique but
        // the compact method works fine and is elegant (and doesn't rely on a 'Random' object).
        // NOTE: DateTime.UtcNow.Ticks is only updated every 15 milliseconds or so, according to some documentation.
    }

    /// <summary>
    /// Name of the application or service - typically the primary namespace or even the assembly
    /// name. Doesn't matter so long as it is Unique, and try not to use ACRONYMS (that's just lazy)
    /// </summary>
    public string ApplicationName { get; set; }

    public string CallerClassName { get; set; }

    /// <summary>
    /// TO ASSIST WITH DEBUGGING.
    /// </summary>
    public string CallerFilePath { get; set; }

    public int CallerLineNumber { get; set; }
    public string CallerMemberName { get; set; }
    public int Day { get; set; }

    /// <summary>
    /// Main text of the event/message, etc. This is designed to be free-form text however... there
    /// is no practical limit to the size of the object so it's possible to store any value here
    /// (even a serialized object if you really want it).
    /// </summary>
    public string EventData { get; set; }

    /// <summary>
    /// From what hosting machine, is the entry from? (makes error tracing a lot easier).
    /// </summary>
    public string MachineName { get; set; }

    public int Month { get; set; }

    /// <summary>
    /// Name of the executing identity (i.e. domain + user name) of the MAIN thread of the running
    /// program. In conjunction with all the other information, this helps identify running processes
    /// and troubleshoot any security context issues.
    /// </summary>
    public string SecurityContext { get; set; }

    public int Year { get; set; }
}