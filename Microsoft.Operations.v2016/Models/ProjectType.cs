using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Microsoft.Operations.TeamFoundationServer
{
    /// <summary>
    /// Psuedo-class for WorkItemTypeCollection
    /// </summary>
    public class ProjectType
    {
        public ProjectType(Project proj)
        {
            this.Name = proj.Name;
            this.WitTypeCollection = proj.WorkItemTypes;
        }

        public string Name { get; set; }

        public WorkItemTypeCollection WitTypeCollection { get; set; }

        public override string ToString()
        {
            return this.Name;
        }
    }
}