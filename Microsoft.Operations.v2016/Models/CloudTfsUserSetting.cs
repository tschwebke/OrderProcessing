using Microsoft.WindowsAzure.Storage.Table;
using System.Data.Services.Common;

namespace Microsoft.Operations
{
    /// <summary>
    /// Use DOMOAIN\alias for RowKey ...
    /// </summary>
    [DataServiceKey("PartitionKey", "RowKey")]
    public class CloudTfsUserSetting : TableEntity
    {
        public CloudTfsUserSetting()
        {
            PartitionKey = "Setting";
        }

        public string SettingName { get; set; }
        public string SettingValue { get; set; }
    }
}