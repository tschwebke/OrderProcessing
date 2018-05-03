using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Microsoft.Operations
{
    /// <summary>
    /// Work item structure which matches the ABO Project template. Actually it is identical to 'CBO
    /// Project' in terms of scoring since they share all the same fields. The item is kept separate
    /// in case any custom logic needs to be applied.
    /// </summary>
    public class ABOProject : XBOCommonFields
    {
        /// <summary>
        /// Looks at the loaded figures and determines if the calculation is correct or not.
        /// </summary>
        public bool CalculationIsCorrect()
        {
            bool correctFigures = true; // start by assuming the calculations are not correct, then we'll prove them otherwise

            if (ImportanceScoreCalculated != ImportanceScore) correctFigures = false;
            if (ComplexityScoreCalculated != ComplexityScore) correctFigures = false;

            // All ABO Projects will use the Composite Score.

            return correctFigures;
        }

        public void Load(WorkItem wi)
        {
            LoadCommonFields(wi);
        }
    }
}