using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Net;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1
    {
        public void CleanupRecords_DoWork()
        {
            //ICredentials credentials = new NetworkCredential(UserIdentity, ServicePassword, "REDMOND");
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(ServerName), credentials);
            tfs.EnsureAuthenticated();
            WorkItemStore workItemStore = new WorkItemStore(tfs);
            List<WorkItem> workitems = new List<WorkItem>();

            string wiql = string.Format(@"SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [System.Tags] FROM WorkItems WHERE [System.TeamProject] = 'CSP'  AND  [System.WorkItemType] = 'SKU Purchase' ORDER BY [System.Id] ");

            workitems = workItemStore.ExecuteQueryText(wiql);

            int count_non_changed_items = 0;

            foreach (WorkItem wi in workitems)
            {
                if (wi.LastChangedAgeInMinutes(60))
                {
                    wi.PartialOpen();

                    wi.CleanSkuInformation(true);

                    var result = wi.Validate();
                    foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result)
                    {
                        Console.WriteLine(wi.Id + info.Status);
                    }

                    if (result.Count == 0 && wi.IsDirty)
                    {
                        // wi.LoadSearchFields(false);
                        // wi.History += "[SYSTEM] Item has been approved by both Field and WOCS. Now awaiting GOC approval (final approval level).";
                        wi.Save();
                    }
                    else
                    {
                        count_non_changed_items++;
                    }
                }
                else
                {
                    count_non_changed_items++;
                }

            }  // end of workitem loop

        }

    }
}