using Microsoft.Operations;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;

/// <summary>
/// Common object for Logging information direct to cloud storage, in the 'cboservicelogs' storage in
/// a really easy-to-use manner. The Azure Table entity structure is optimized for deleting old logs,
/// and casual veiwing.
/// </summary>
/// <example>AzureLogger Log = new AzureLogger(); Log.WriteEntry("blah")</example>
/// <todo>Needs Async calling ... https://msdn.microsoft.com/en-us/library/azure/dn435331.aspx http://blog.stephencleary.com/2012/02/async-and-await.html</todo>
public class AzureLogger
{
    private CloudTable _activity_store;
    private string _assembly_name = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
    private string _executing_identity = (System.Environment.UserDomainName.ToUpper() + @"\" + System.Environment.UserName).ToLower();
    private string _machine_name = System.Environment.MachineName.ToUpper();
    private CloudStorageAccount _storage_account;

    /// <summary>
    /// Create new component for easy reference to your Azure objects. It will use the name for the
    /// Partition Key in the Azure table. Recommend creating only once when program starts, then
    /// re-using. As far as we know, the object does not keep connections open, that sort of thing.
    /// </summary>
    public AzureLogger()
    {
        _storage_account = CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=cboservicelogs;AccountKey=7psl9DFh8F3AISj1FoqwtMyBNQm4ni7r2ES1tF5PiEsJgxDbYzZ7LFDCLQHBU0Hmv2NN+we2WmJfbTP12VxClA==");
        _activity_store = _storage_account.CreateCloudTableClient().GetTableReference("ApplicationChatter");
        WriteLine(string.Format("VERSION: {0} PROCESS STARTED", Assembly.GetExecutingAssembly().GetVersion(), _machine_name));
    }

    /// <summary>
    /// Writes supplied text to the Azure Tables log. All the other attributes of the entry are
    /// discovered when the 'AzureLogger' object is first created. Due to using the optional
    /// parameters for determining caller information; we have sacrified the ability to use
    /// arguments/parameters - the two constructs are incompatible with each other.
    /// http://stackoverflow.com/questions/14505573/optional-parameters-together-with-an-array-of-parameters
    /// Also chose not to walk the stack, since the 'callerinfo' values have virtually no performance
    /// penalty. http://blog.slaks.net/2011/10/caller-info-attributes-vs-stack-walking.html
    /// </summary>
    /// <remarks>
    /// It is possible to override the default values, e.g. if you wish to supply the method name (or
    /// something custom)
    /// </remarks>
    public void WriteLine(string message, [CallerMemberName] string methodName = "", [CallerFilePath] string sourceFilePath = "", [CallerLineNumber] int sourceLineNumber = 0)
    {
        LogDetail item = new LogDetail(_assembly_name, methodName.Replace(".ctor", "{Constructor}"), Path.GetFileName(sourceFilePath), sourceLineNumber);
        item.EventData = message;
        item.MachineName = _machine_name;
        item.SecurityContext = _executing_identity;

        TableOperation update = TableOperation.InsertOrReplace(item); // in case milliseconds collision. very unlikely but this is an easier handler
        _activity_store.Execute(update);
    }
}