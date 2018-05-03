using Microsoft.WindowsAzure.Storage.Table;
using System.Data.Services.Common;

namespace Microsoft.Operations
{
    /// <summary>
    /// Use display name for RowKey ...
    /// </summary>
    [DataServiceKey("PartitionKey", "RowKey")]
    public class CloudContact : TableEntity
    {
        public CloudContact()
        {
            PartitionKey = "Contact";
        }

        public string Alias { get; set; }
        public string DisplayName { get; set; }
        public double ImportanceScore { get; set; }
        public string InitiativeRowKey { get; set; }
        public int ProjectID { get; set; }
        public string ProjectStatus { get; set; }

        // these are used for filtering and grouping for display
        public string ProjectTitle { get; set; }

        public string Role { get; set; }
        public string RoleCode { get; set; }
    }
}