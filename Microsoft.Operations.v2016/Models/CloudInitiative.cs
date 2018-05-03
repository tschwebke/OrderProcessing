using Microsoft.WindowsAzure.Storage.Table;
using System.Collections.Generic;
using System.Data.Services.Common;

namespace Microsoft.Operations
{
    [DataServiceKey("PartitionKey", "RowKey")]
    public class CloudInitiative : TableEntity
    {
        public CloudInitiative()
        {
            TeamOwnershipAllTeams = new List<string>();
            PartitionKey = "Summary";
            ImportanceScore = 0;
        }

        /// <summary>
        /// Large string which can contain summaries useful for filtering (general use)
        /// </summary>
        public string FilterValues { get; set; }

        /// <summary>
        /// This is actually the CUMULATIVE score of all sub-projects importance, which is then used
        /// later to determine rank.
        /// </summary>
        public double ImportanceScore { get; set; }

        public string InitiativeCode { get; set; }

        /// <summary>
        /// This is a red/yellow/green, but allows multiples, e.g. RRRGGGYY
        /// </summary>
        public string ProjectStatus { get; set; }

        public int Rank { get; set; }

        public string Strategy { get; set; }

        /// <summary>
        /// Which OBO (Mike Novasio) strategy is paired with this initiative? (if any?) Abbreviated
        /// form, 1-1 relationship
        /// </summary>
        public string StrategyCode { get; set; }

        public string TeamOwnership { get; set; }

        public List<string> TeamOwnershipAllTeams { get; set; }

        /// <summary>
        /// Primary text name of the intiative (display version)
        /// </summary>
        public string Title { get; set; }
    }
}