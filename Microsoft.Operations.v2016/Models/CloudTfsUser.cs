using Microsoft.WindowsAzure.Storage.Table;
using System.Data.Services.Common;

namespace Microsoft.Operations
{
    /// <summary>
    /// Use DOMOAIN\alias for RowKey ...
    /// </summary>
    [DataServiceKey("PartitionKey", "RowKey")]
    public class CloudTfsUser : TableEntity
    {
        public CloudTfsUser()
        {
            PartitionKey = "User";
        }

        public CloudTfsUser(string rowKey)
        {
            ETag = "*";
            PartitionKey = "User";
            RowKey = rowKey;
        }

        public string DisplayName { get; set; }
        public string GroupMembership { get; set; }
    }
}