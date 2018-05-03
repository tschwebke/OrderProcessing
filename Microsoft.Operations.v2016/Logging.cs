using Microsoft.Operations;
using Microsoft.WindowsAzure.Storage.Table;
using System.Diagnostics;
using System.Reflection;

/// <summary>
/// Cloud-based logging components
/// </summary>
public static class Logging
{
    /// <summary>
    /// Shortcut, will write as 'Information' message to the specified log.
    /// </summary>
    public static void WriteEntry(this CloudTable Table, string message)
    {
        Table.WriteEntry(EventLogEntryType.Information, message);
    }

    /// <summary>
    /// Shortcut, will write as 'Information' message to the specified log, allows parameters.
    /// </summary>
    public static void WriteEntry(this CloudTable Table, string message, params object[] args)
    {
        Table.WriteEntry(EventLogEntryType.Information, message, args);
    }

    /// <summary>
    /// Full 'write entry' which accepts arguments.
    /// Note: Partition key will be the calling assembly,
    /// </summary>
    public static void WriteEntry(this CloudTable Table, EventLogEntryType level, string message, params object[] args)
    {
        string natureOfEvent = "Unclassified";
        switch (level)
        {
            case EventLogEntryType.Information: natureOfEvent = "Information"; break;
            case EventLogEntryType.Warning: natureOfEvent = "Warning"; break;
            case EventLogEntryType.Error: natureOfEvent = "Error"; break;
        }

        Assembly from = Assembly.GetCallingAssembly();
        LogEntry item = new LogEntry(from.FullName.Substring(0, from.FullName.IndexOf(','))) { Level = natureOfEvent, EventData = message.Fill(args) };
        TableOperation update = TableOperation.InsertOrReplace(item); // in case milliseconds collision
        Table.Execute(update);
    }

    // TODO: I would like to integate the calling method into the name automatically, instead of
    // System.Reflection.MethodBase.GetCurrentMethod().Name HOWEVER, this isn't safe in production
    // code (release) ... only Debug build :(

    // Current class name :

    //this.GetType().Name;

    // Current method name:

    //using System.Reflection;

    //MethodBase.GetCurrentMethod().Name;

    // the calling method:

    // StackTrace stackTrace = new StackTrace(); StackFrame stackFrame = stackTrace.GetFrame(1);
    // MethodBase methodBase = stackFrame.GetMethod();
}