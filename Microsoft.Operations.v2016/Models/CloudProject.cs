using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Data.Services.Common;

namespace Microsoft.Operations
{
    /// <summary>
    /// Fields common to CBO template (at this stage)
    /// </summary>
    [DataServiceKey("PartitionKey", "RowKey")]
    public class CloudProject : TableEntity
    {
        public CloudProject()
        {
            PartitionKey = "Project";
        }

        public string AreaPath { get; set; }

        /// <summary>
        /// "blank, partner, customer, developer, internal" From the multi-select option. Used for
        /// filtering in the website.
        /// </summary>
        public string Audience { get; set; }

        public string Author { get; set; }
        public string BenefitScore { get; set; }
        public string Blocker { get; set; }
        public string BlockerDetail { get; set; }

        /// <summary>
        /// Multi-select fields, contains delimited []
        /// </summary>
        public string CapabilitiesImpacted { get; set; }

        public string ChangedBy { get; set; }
        public DateTime ChangedDate { get; set; }
        public string CompositeScore { get; set; }

        /// <summary>
        /// combination of the two
        /// </summary>
        public string CompoundKey { get; set; }

        public string DateBRD { get; set; }
        public string DateBusinessGroupGOLIVE { get; set; }
        public string DateCheckin { get; set; }
        public string DateEngineeringGOLIVE { get; set; }
        public string DateFinishTrainingAndReadiness { get; set; }
        public string DateFinishUAT { get; set; }
        public string DateKickoff { get; set; }
        public string DateLaunch { get; set; }
        public string DateLaunchActual { get; set; }

        // common value ...
        public string DateLaunchTarget { get; set; }

        public string DateStartTrainingAndReadiness { get; set; }
        public string DateStartUAT { get; set; }
        public string Description { get; set; }

        // ABO Specific items that we hope to make common one day!
        public string EffortScore { get; set; }

        public string EngineeringDependencies { get; set; }
        public string FinishDate { get; set; }
        public string HelpWanted { get; set; }
        public string Highlights { get; set; }
        public string HighlightsDateChanged { get; set; }
        public int Id { get; set; }
        public string ImpactDocuments { get; set; }
        public string ImpactRegions { get; set; }
        public string ImpactSystems { get; set; }
        public string ImpactTeams { get; set; }

        /// <summary>
        /// Value from the common scoring system. This will be summed in order to get the most
        /// important initiative.
        /// </summary>
        public double ImportanceScore { get; set; }

        /// <summary>
        /// 2. The 'set' to which this project belongs to.
        /// </summary>
        public string Initiative { get; set; }

        /// <summary>
        /// this is designed to fully replace 'InitiativeRowKey' (is the same value)
        /// </summary>
        public string InitiativeCode { get; set; }

        /// <summary>
        /// Depracated, will replace.
        /// </summary>
        public string InitiativeRowKey { get; set; }

        public string Lowlights { get; set; }
        public string LowlightsDateChanged { get; set; }

        // important one important one
        public string LTOwner { get; set; }

        public string NodeName { get; set; }

        /// <summary>
        /// This is an illegitimate field -&gt; shouldn't exist, this is a subclass of 'dependencies'
        /// in general
        /// </summary>
        public string OperatingModelsImpacted { get; set; }

        public string PriorityCompete { get; set; }
        public string PriorityCompeteComments { get; set; }
        public string PriorityCompliance { get; set; }
        public string PriorityComplianceComments { get; set; }
        public string PriorityConstraints { get; set; }
        public string PriorityConstraintsComments { get; set; }
        public string PriorityCost { get; set; }
        public string PriorityCostComments { get; set; }
        public string PriorityCPE { get; set; }
        public string PriorityCPEComments { get; set; }
        public string ProblemStatement { get; set; }
        public string ProjectOwner { get; set; }

        /// <summary>
        /// Project priority is an Integer field inside the system. It is not always used for
        /// calculations, but is nearly always present regardless of template.
        /// </summary>
        public string ProjectPriority { get; set; }

        public string ProjectStatus { get; set; }
        public string ProjectStatusDateChanged { get; set; }
        public string ProjectStatusExplanation { get; set; }
        public string ProjectStatusExplanationDateChanged { get; set; }
        public string ProjectStatusPrevious { get; set; }
        public int Rank { get; set; }
        public string Recommendation { get; set; }
        public string Requestor { get; set; }
        public string ResourceCapabilityManager { get; set; }

        public string ResourceConsultant { get; set; }

        public string ResourceDeployManager { get; set; }

        public string ResourceExecutionManager { get; set; }

        public string ResourcePlanningManager { get; set; }

        /// <summary>
        /// Contains tabular data, that we have to unpack. List of People/folk ... and their alias
        /// </summary>
        public string ResourcesListingDelimited { get; set; }

        public string ResourceTestManager { get; set; }

        public string ResourceTrainingReadinessSupport { get; set; }

        public string ScoringSubmissionCount { get; set; }

        public string StartDate { get; set; }

        public string State { get; set; }

        /// <summary>
        /// 1. This is the 'Mike Novasio' strategy level.
        /// </summary>
        public string Strategy { get; set; }

        /// <summary>
        /// and the abbreviated, lower-case value, for use with the website ...
        /// </summary>
        public string StrategyCode { get; set; }

        public string StrategyMSOPS { get; set; }

        public string StrategyOBO { get; set; }

        public string SubState { get; set; }

        public string SuccessMeasurement { get; set; }

        public string SuccessMeasurementTarget { get; set; }

        /// <summary>
        /// Does this belong to ABO, CBO or SBO? (singular only) (determined from the original name
        /// of template)
        /// </summary>
        public string TeamOwnership { get; set; }

        public string Title { get; set; }

        /// <summary>
        /// This is actually a more compact name for 'Critical Issues, Risks and Mitigation Plans'
        /// </summary>
        public string Updates { get; set; }

        /// <summary>
        /// The URL as given for a link back to the main TFS site.
        /// </summary>
        public string VSTFUrl { get; set; }

        public string WorkItemType { get; set; }
    }
}