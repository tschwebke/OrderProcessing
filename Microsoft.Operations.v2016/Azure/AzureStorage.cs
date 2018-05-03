using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.WindowsAzure.Storage.Table.DataServices;

/// <summary>
/// Common object for accessing cloud storage data in the CBO subscription. Contains shortcuts which
/// makes it easier to manage objects. Handles one table and one queue at a time
/// </summary>
/// <example>
/// AzureStorageAccount Azure = new AzureStorageAccount("cboservicelogs"); Azure.LoadQueue("Stuff"); Azure.Queue.AddMessage(CloudQueueMessageExtensions.Serialize(s));
/// </example>
public class AzureStorageAccount
{
    private static CloudQueue _queue_ref;
    private static CloudQueueClient _queueClient;
    private static CloudStorageAccount _storage_account;
    private static CloudTableClient _table_client;
    private static CloudTable _table_ref;
    private string AccountKey = string.Empty;

    /// <summary>
    /// Create new component for easy reference to your Azure objcts
    /// </summary>
    /// <param name="container">Name of the storage account you want to connect to</param>
    public AzureStorageAccount(string container) : this(container, string.Empty)
    {
    }

    /// <summary>
    /// </summary>
    /// <param name="container"></param>
    /// <param name="tableName">
    /// Optionally load the table name direct from the constructor (as a shortcut ... one less line
    /// of code if you're only working with one table)
    /// </param>
    public AzureStorageAccount(string container, string tableName)
    {
        switch (container)
        {
            case "cboalerts": AccountKey = "H74ut2+bhZ8563iBB2ri46mdqvjRkZDT8beNGODDdrUZrGeRFg4ex5U4i37WnETyh8APgHoXNPFhG5o0xLMIVw=="; break;
            case "cboservicelogs": AccountKey = "7psl9DFh8F3AISj1FoqwtMyBNQm4ni7r2ES1tF5PiEsJgxDbYzZ7LFDCLQHBU0Hmv2NN+we2WmJfbTP12VxClA=="; break;
            case "cbovstftxfr": AccountKey = "tTZIfzMtSMa64PsCPl9iI7NUDwvkttkCmajnGGClASJr4AHY3g5V5m0TA8Kgt/IMSi+lWqmOQujJFa3fbl7M4A=="; break;
            case "opsvstf": AccountKey = "TomLlgPtRfip3r6sQWcm7UxBZI6PeWm1lhUky4rD0QGG6cp9pVnlb9/Dne3C8FrXGArl0Rog4iBicZxp/1WsMA=="; break;
        }

        _storage_account = CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=" + container + ";AccountKey=" + AccountKey);
        _table_client = _storage_account.CreateCloudTableClient();
        _queueClient = _storage_account.CreateCloudQueueClient();

        if (!string.IsNullOrEmpty(tableName)) { LoadTable(tableName); }
    }

    /// <summary>
    /// Generic queue
    /// </summary>
    public CloudQueue Queue
    {
        get { return _queue_ref; }
        set { _queue_ref = value; }
    }

    /// <summary>
    /// The table object which was previously loaded.
    /// </summary>
    public CloudTable Table
    {
        get { return _table_ref; }
        set { _table_ref = value; }
    }

    /// <summary>
    /// Reference to the client object created when a table was loaded.
    /// </summary>
    public CloudTableClient TableClient
    {
        get { return _table_client; }
    }

    /// <summary>
    /// Use this endpoint to do unstructured queries.
    /// </summary>
    /// <example>
    /// var query = from entity in AzureScoringData.TableContext.CreateQuery ABOScore
    /// ("AdvertisingBusinessOperations") where entity.WebsiteIsModified.Equals(true) select entity;
    /// </example>
    public TableServiceContext TableContext
    {
        get { return _table_ref.ServiceClient.GetTableServiceContext(); }
    }

    public void LoadQueue(string queueName)
    {
        _queue_ref = _queueClient.GetQueueReference(queueName);
    }

    public void LoadTable(string tableName)
    {
        _table_ref = _table_client.GetTableReference(tableName);
    }
}