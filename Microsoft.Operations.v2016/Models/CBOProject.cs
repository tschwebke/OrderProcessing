using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Microsoft.Operations
{
    /// <summary>
    /// Work item strcture which matches the CBO Project template.
    /// </summary>
    public class CBOProject : XBOCommonFields
    {
        /// <summary>
        /// Looks at the loaded figures and determines if the calculation is correct or not.
        /// </summary>
        public bool CalculationIsCorrect()
        {
            bool correctFigures = true; // start by assuming the calculations are not correct, then we'll prove them otherwise

            // if (BusinessImpactMixCalculated != BusinessImpactMix) correctFigures = false; if
            // (BusinessImpactScoreCalculated != BusinessImpactScore) correctFigures = false;
            if (ImportanceScoreCalculated != ImportanceScore) correctFigures = false;
            if (ComplexityScoreCalculated != ComplexityScore) correctFigures = false;

            // new rule introduced in V20140804, Every CBO Project item ought to have a non-root value.
            if (AreaPath == "Online Operations") correctFigures = false;

            return correctFigures;
        }

        public void Load(WorkItem wi)
        {
            LoadCommonFields(wi);
        }
    }
}