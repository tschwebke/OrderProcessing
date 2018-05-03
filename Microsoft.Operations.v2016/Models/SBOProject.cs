using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Microsoft.Operations
{
    /// <summary>
    /// Work item strcture which matches the SBO Project template. This is a simpler version compared
    /// with CBO or SBO.
    /// </summary>
    public class SBOProject : XBOCommonFields
    {
        /// <summary>
        /// Looks at the loaded figures and determines if the calculation is correct or not.
        /// </summary>
        public bool CalculationIsCorrect()
        {
            bool correctFigures = true; // start by assuming the calculations are not correct, then we'll prove them otherwise

            // if (BusinessImpactMixCalculated != BusinessImpactMix) correctFigures = false; if
            // (BusinessImpactScoreCalculated != BusinessImpactScore) correctFigures = false; if
            // (ImportanceScoreCalculated != ImportanceScore) correctFigures = false; if
            // (ComplexityScoreCalculated != ComplexityScore) correctFigures = false;

            // Every SBO Project item ought to have a non-root value.
            if (AreaPath == "Online Operations") correctFigures = false;

            return correctFigures;
        }

        public void Load(WorkItem wi)
        {
            LoadCommonFields(wi);
        }
    }
}