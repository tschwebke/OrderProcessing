using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;

namespace Microsoft.Operations
{
    public static partial class TeamFoundationServerHelper
    {
        /// <summary>
        /// Note: This version does not support wildcards like @project - could be extended.
        /// </summary>
        public static List<WorkItem> ExecuteQueryText(this WorkItemStore workItemStore, string queryTextIn)
        {
            // The line below allows you to use
            queryTextIn = queryTextIn.Replace("&lt;", "<").Replace("&gt;", ">");

            var q = new Query(workItemStore, queryTextIn);

            List<WorkItem> wis = new List<WorkItem>();

            if (q.IsLinkQuery)
            {
                var queryResults = q.RunLinkQuery();
                foreach (WorkItemLinkInfo item in queryResults)
                {
                    wis.Add(workItemStore.GetWorkItem(item.TargetId));
                }
            }
            else
            {
                var queryResults = q.RunQuery();
                foreach (WorkItem item in queryResults)
                {
                    wis.Add(item);
                }
            }
            return wis;
        }
    }
}