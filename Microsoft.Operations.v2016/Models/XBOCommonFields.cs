using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;

namespace Microsoft.Operations
{
    /// <summary>
    /// Fields common to CBO template (at this stage)
    /// </summary>
    public class XBOCommonFields : WorkItemSystemFields
    {
        public string ProblemStatement;
        public string Recommendation;

        // added 2015 to support ABO Scoring website
        public string RiskIfNotImplemented;

        private decimal _complexityProjectSizeScore, _complexityCapabilitiesScore, _complexityScopeScore;
        private decimal _importanceComplianceScore, _importanceCPEScore, _importanceCostScore, _importanceCompeteScore, _importanceConstraintsScore;

        private enum ScoringAspect
        {
            Importance,
            Complexity
        }

        public string AreaPath { get; set; }
        public string Author { get; set; }

        /// <summary>
        /// Value between 1 and 225, calculated from various "Importance" and "Complexity" fields.
        /// </summary>
        public decimal BusinessImpactScore { get; set; }

        public decimal BusinessImpactScoreCalculated { get; set; }
        public string ComplexityCapabilities { get; set; }

        /// <summary>
        /// String value, sizes from "Microsoft.Operations.Project.Complexity.Size" values are: L, M,
        /// S, XL These are then used for some calculations based on the value ...
        /// </summary>
        public string ComplexityProjectSize { get; set; }

        public string ComplexityScope { get; set; }
        public decimal ComplexityScore { get; set; }
        public decimal ComplexityScoreCalculated { get; set; }
        public string[] ImportanceCompeteText { get; set; }
        public string[] ImportanceComplexityText { get; set; }

        /// <summary>
        /// Importance values are: 'Extremely High' = 5 'High' = 3 'Medium' = 1 'Low' = 0
        /// </summary>
        public string[] ImportanceComplianceText { get; set; }

        public string[] ImportanceConstraintsText { get; set; }
        public string[] ImportanceCostText { get; set; }

        // added 2015 to support ABO Scoring website added 2015 to support ABO Scoring website
        public string[] ImportanceCPEText { get; set; }

        // also additional .. CBO Project doesn't use this one in the layout, however it's in the
        // template field listing for overally compatibility
        public decimal ImportanceScore { get; set; }

        public decimal ImportanceScoreCalculated { get; set; }
        public string NodeName { get; set; }

        /// <summary>
        /// Project priority is an Integer field inside the system. It is not always used for
        /// calculations, but is nearly always present regardless of template.
        /// </summary>
        public decimal ProjectPriority { get; set; }

        public string ProjectSize { get; set; }
        public string Requestor { get; set; }

        /// <summary>
        /// Typically called after all the values are loaded, or independently for a recalculation.
        /// Supply value: ZERO = calculate against the loaded ones (normal CBO Project) ONE =
        /// calculate against other calculated fields (e.g. for ABO Search)
        /// </summary>
        public void Calculate(int arrayPosition)
        {
            _complexityProjectSizeScore = CalculateTextEquivalent(ScoringAspect.Complexity, ComplexityProjectSize);
            _complexityCapabilitiesScore = CalculateTextEquivalent(ScoringAspect.Complexity, ComplexityCapabilities);
            _complexityScopeScore = CalculateTextEquivalent(ScoringAspect.Complexity, ComplexityScope);

            _importanceComplianceScore = CalculateTextEquivalent(ScoringAspect.Importance, ImportanceComplianceText[arrayPosition]);
            _importanceCPEScore = CalculateTextEquivalent(ScoringAspect.Importance, ImportanceCPEText[arrayPosition]);
            _importanceCostScore = CalculateTextEquivalent(ScoringAspect.Importance, ImportanceCostText[arrayPosition]);
            _importanceCompeteScore = CalculateTextEquivalent(ScoringAspect.Importance, ImportanceCompeteText[arrayPosition]);
            _importanceConstraintsScore = CalculateTextEquivalent(ScoringAspect.Importance, ImportanceConstraintsText[arrayPosition]);

            ComplexityScoreCalculated = _complexityProjectSizeScore + _complexityCapabilitiesScore + _complexityScopeScore;
            // ImportanceScoreCalculated = _importanceCostScore + _importanceComplianceScore +
            // _importanceCompeteScore + _importanceCPEScore + _importanceConstraintsScore; // v20140530
            ImportanceScoreCalculated = Math.Round((_importanceCostScore * 1) + (_importanceComplianceScore * new decimal(1.25)) + (_importanceCompeteScore * 2) + (_importanceCPEScore * new decimal(1.5)) + (_importanceConstraintsScore * 1), 0); // v20140615
            //BusinessImpactScoreCalculated = ComplexityScore * ImportanceScore;

            // tiered approach for scoring.
            //if(BusinessImpactScoreCalculated < 10) { BusinessImpactMixCalculated = "Get it Done"; }
            //else if (BusinessImpactScoreCalculated < 30) { BusinessImpactMixCalculated = "We Need it"; }
            //else if (BusinessImpactScoreCalculated < 91) { BusinessImpactMixCalculated = "Differentiator"; }
            //else if (BusinessImpactScoreCalculated < 157) { BusinessImpactMixCalculated = "Game Changer"; }
            //else if (BusinessImpactScoreCalculated < 226) { BusinessImpactMixCalculated = "Epic Stuff"; }
            //else { BusinessImpactMixCalculated = "Epic x 2"; } // currently there's no value specified for anything higher
        }

        //public string BusinessImpactMix { get; set; }
        //public string BusinessImpactMixCalculated { get; set; }
        public void LoadCommonFields(WorkItem wi)
        {
            Id = wi.Id;
            State = wi.State;
            Title = wi.Title;
            AreaPath = wi.AreaPath;
            ChangedDate = wi.ChangedDate;
            ChangedBy = wi.ChangedBy;

            ImportanceComplianceText = new string[2];
            ImportanceCPEText = new string[2];
            ImportanceCostText = new string[2];
            ImportanceCompeteText = new string[2];
            ImportanceConstraintsText = new string[2];
            ImportanceComplexityText = new string[2];

            ProblemStatement = wi.GetFieldValue<string>("Microsoft.Operations.WorkItem.Description.Problem");
            RiskIfNotImplemented = wi.GetFieldValue<string>("Microsoft.Operations.WorkItem.Description.Threat");
            Recommendation = wi.GetFieldValue<string>("Microsoft.Operations.WorkItem.Description.Proposal");
            Author = wi.GetFieldValue<string>("Microsoft.Operations.WorkItem.Ownership.Author");
            Requestor = wi.GetFieldValue<string>("Microsoft.Operations.WorkItem.Ownership.Requestor");

            if (wi.Fields["System.NodeName"] != null) NodeName = wi.Fields["System.NodeName"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Complexity.Size"] != null) ComplexityProjectSize = wi.Fields["Microsoft.Operations.Project.Complexity.Size"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Complexity.Capabilities"] != null) ComplexityCapabilities = wi.Fields["Microsoft.Operations.Project.Complexity.Capabilities"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Complexity.Scope"] != null) ComplexityScope = wi.Fields["Microsoft.Operations.Project.Complexity.Scope"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.Compliance.Actual"] != null) ImportanceComplianceText[0] = wi.Fields["Microsoft.Operations.Project.Impact.Compliance.Actual"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.CPE.Actual"] != null) ImportanceCPEText[0] = wi.Fields["Microsoft.Operations.Project.Impact.CPE.Actual"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.Cost.Actual"] != null) ImportanceCostText[0] = wi.Fields["Microsoft.Operations.Project.Impact.Cost.Actual"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.Compete.Actual"] != null) ImportanceCompeteText[0] = wi.Fields["Microsoft.Operations.Project.Impact.Compete.Actual"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.Constraints.Actual"] != null) ImportanceConstraintsText[0] = wi.Fields["Microsoft.Operations.Project.Impact.Constraints.Actual"].Value.ToString();
            if (wi.Fields["Microsoft.Operations.Project.Impact.Complexity.Actual"] != null) ImportanceComplexityText[0] = wi.Fields["Microsoft.Operations.Project.Impact.Complexity.Actual"].Value.ToString();

            // Now load the main scoring values, if applicable

            if (wi.Fields["Microsoft.Operations.Project.Score.Complexity"] != null) ComplexityScore = Convert.ToDecimal(wi.Fields["Microsoft.Operations.Project.Score.Complexity"].Value);
            if (wi.Fields["Microsoft.Operations.Project.Score.Importance"] != null) ImportanceScore = Convert.ToDecimal(wi.Fields["Microsoft.Operations.Project.Score.Importance"].Value);
            // if (wi.Fields["Microsoft.Operations.Project.Complexity.Score"] != null)
            // BusinessImpactScore =
            // Convert.ToDecimal(wi.Fields["Microsoft.Operations.Project.Complexity.Score"].Value);
            // if (wi.Fields["Microsoft.Operations.Project.Complexity.Mix"] != null)
            // BusinessImpactMix =
            // wi.Fields["Microsoft.Operations.Project.Complexity.Mix"].Value.ToString(); // obsolete
            // as per v20140615

            Calculate(0);
        }

        private decimal CalculateTextEquivalent(ScoringAspect input, string value)
        {
            switch (value)
            {
                // case "Extremely High": return 5;
                case "High": return 5;
                case "Medium": return 3;
                case "Low": return 1;

                case "XL": return 5;
                case "L": return 3;
                case "M": return 1;
                case "S": return 0;
                default: return 0;
            }
        }
    }
}