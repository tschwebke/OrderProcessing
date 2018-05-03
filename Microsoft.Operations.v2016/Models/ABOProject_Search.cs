using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;

namespace Microsoft.Operations
{
    /// <summary>
    /// These are the 'team values' which are tracked separately in the 'Search template' grid. The
    /// numeric values can be used in the decimal array, be careful with the ordering !
    /// </summary>
    public enum ExternalTeam
    {
        /// <summary>
        /// Special, holds the average for all groups
        /// </summary>
        Average = 0,

        GSDS = 1,
        AdSupport = 2,
        CSS = 3,
        Finance = 4,
        WOCs = 5,
        AdOps = 6,
        Sales = 7, // formerly 'TQ'
        Accounting = 8
    }

    /// <summary>
    /// Similar to ABO Project, but the 'Search' version has so many different fields that it really
    /// is like a completely separate template rather than being a branch of "ABO Project". Isolated
    /// now as a separate class to make future adjustments easier.
    ///
    /// There are over 100 fields involved in their calculation mechanic so we'll use shorthand where
    /// possible, including arrays for similar grid values.
    /// </summary>
    public class ABOProject_Search : XBOCommonFields
    {
        // TWO dimensional array. Position ZERO holds loaded values (i.e. from the work item)
        // Position ONE holds calculated values (i.e. based on what gets calculated)
        // NOTE: NULL values from TFS are stored as "-1" values

        public decimal[,] BenefitComplianceRisk;
        public decimal[,] BenefitCustomerSatisfaction;
        public decimal[,] BenefitEffectiveness;
        public decimal[,] BenefitEfficiency;
        public decimal[,] BenefitEmployeeSatisfaction;
        public decimal[,] BenefitReach;
        public decimal[,] BenefitScale;
        public string BusinessCaseJustification;
        public string Comments;
        public decimal[,] EffortEngagement;
        public decimal[,] EffortExpense;
        public decimal[,] EffortGlobalCoordination;
        public decimal[,] EffortGroupCoordination;
        public string[] EffortLevelAdOps = new string[2];
        public decimal[,] EffortPlatformCoordination;

        public double[] EstimatedHours = new double[2];
        public decimal[] FMXTotalScore = new decimal[2];
        public string[] FMXUsage = new string[2];
        public decimal[,] ImportanceCompete;
        public decimal[,] ImportanceComplexity;
        public decimal[,] ImportanceCompliance;
        public decimal[,] ImportanceConstraints;
        public decimal[,] ImportanceCost;
        public decimal[,] ImportanceCPE;
        public string InformationOnMandatoryForComplianceLegalRegulatory;
        public string MandatoryForComplianceLegalRegulatory;
        public string OperationsBenefits;

        // additional, ABO Search specific field (none of the other CBO orgs use this one)
        public int[] Priority = new int[2];

        // added 2015 to support ABO Scoring website
        public string RevenueImpact;

        public decimal[,] ScoreBenefits;
        public decimal[,] ScoreComposite;
        public decimal[,] ScoreEfforts;
        public int[] TotalTeamSubmissions = new int[2];
        // new in Nov 2014. ref. Task new in Nov 2014. new in Nov 2014.

        // Some of these are very similar to: ABOIssue, but differ by data type.

        // added 2015 to support ABO Scoring website added 2015 to support ABO Scoring website added
        // 2015 to support ABO Scoring website added 2015 to support ABO Scoring website this may
        // possibly be later, a common field (it should be)

        /// <summary>
        /// Looks at the loaded figures and determines if the calculation is correct or not.
        /// </summary>
        public bool CalculationIsCorrect()
        {
            bool correctFigures = true; // start by assuming the calculations are not correct, then we'll prove them otherwise

            // Task: compare the array loaded values (position zero) with the calculated values
            //       (position one) For each one,

            // NOTE: These only apply when 'Online Operations\Advertising\Search' is selected in the
            //       Area Path.

            if (!LoadedEqualsCalculated(BenefitComplianceRisk)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitEfficiency)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitEffectiveness)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitScale)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitReach)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitCustomerSatisfaction)) correctFigures = false;
            if (!LoadedEqualsCalculated(BenefitEmployeeSatisfaction)) correctFigures = false;
            if (!LoadedEqualsCalculated(EffortExpense)) correctFigures = false;
            if (!LoadedEqualsCalculated(EffortEngagement)) correctFigures = false;
            if (!LoadedEqualsCalculated(EffortGroupCoordination)) correctFigures = false;
            if (!LoadedEqualsCalculated(EffortGlobalCoordination)) correctFigures = false;
            if (!LoadedEqualsCalculated(EffortPlatformCoordination)) correctFigures = false;

            if (!LoadedEqualsCalculated(ScoreBenefits)) correctFigures = false;
            if (!LoadedEqualsCalculated(ScoreEfforts)) correctFigures = false;
            if (!LoadedEqualsCalculated(ScoreComposite)) correctFigures = false;

            if (!LoadedEqualsCalculated(ImportanceCompliance)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCPE)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCost)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCompete)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceConstraints)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceComplexity)) correctFigures = false;

            if (!LoadedEqualsCalculated(ImportanceComplianceText)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCPEText)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCostText)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceCompeteText)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceConstraintsText)) correctFigures = false;
            if (!LoadedEqualsCalculated(ImportanceComplexityText)) correctFigures = false;

            if (ImportanceScoreCalculated != ImportanceScore) correctFigures = false;
            if (ComplexityScoreCalculated != ComplexityScore) correctFigures = false;

            if (TotalTeamSubmissions[0] != TotalTeamSubmissions[1]) correctFigures = false;
            if (FMXTotalScore[0] != FMXTotalScore[1]) correctFigures = false;

            if (Priority[0] != Priority[1]) correctFigures = false;
            if (EffortLevelAdOps[0] != EffortLevelAdOps[1]) correctFigures = false;
            if (EstimatedHours[0] != EstimatedHours[1]) correctFigures = false;

            // new fields as well.

            return correctFigures;
        }

        public void Load(WorkItem wi)
        {
            LoadCommonFields(wi);

            BusinessCaseJustification = wi.GetFieldValue<string>("Microsoft.Operations.Requirement.Justification.Full");
            RevenueImpact = wi.GetFieldValue<string>("Microsoft.Operations.Project.Impact.Revenue.Writeup");
            OperationsBenefits = wi.GetFieldValue<string>("Microsoft.Operations.Project.Benefits.Operations");
            MandatoryForComplianceLegalRegulatory = wi.GetFieldValue<string>("Microsoft.Operations.Project.Regulation");
            InformationOnMandatoryForComplianceLegalRegulatory = wi.GetFieldValue<string>("Microsoft.Operations.Project.Regulation.Comments");
            Comments = wi.GetFieldValue<string>("Microsoft.VSTS.CMMI.Comments");

            BenefitComplianceRisk = new decimal[9, 2];
            BenefitEfficiency = new decimal[9, 2];
            BenefitEffectiveness = new decimal[9, 2];
            BenefitScale = new decimal[9, 2];
            BenefitReach = new decimal[9, 2];
            BenefitCustomerSatisfaction = new decimal[9, 2];
            BenefitEmployeeSatisfaction = new decimal[9, 2];
            EffortExpense = new decimal[9, 2];
            EffortEngagement = new decimal[9, 2];
            EffortGroupCoordination = new decimal[9, 2];
            EffortGlobalCoordination = new decimal[9, 2];
            EffortPlatformCoordination = new decimal[9, 2];

            ScoreBenefits = new decimal[9, 2];
            ScoreEfforts = new decimal[9, 2];
            ScoreComposite = new decimal[9, 2];

            ImportanceCompliance = new decimal[9, 2];
            ImportanceCPE = new decimal[9, 2];
            ImportanceCost = new decimal[9, 2];
            ImportanceCompete = new decimal[9, 2];
            ImportanceConstraints = new decimal[9, 2];
            ImportanceComplexity = new decimal[9, 2]; // new

            // now load fields which are specific to this template. Please note! the field must be
            // present inside the template, otherwise the fetch process will fail ! If there is an
            // error with calculating scores, and it is known that the work item template has been
            // modified, check here first.

            // Averages first ...

            BenefitComplianceRisk[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk");
            BenefitEfficiency[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency");
            BenefitEffectiveness[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness");
            BenefitScale[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale");
            BenefitReach[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach");
            BenefitCustomerSatisfaction[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Satisfaction.Customer"); // slightly different name
            BenefitEmployeeSatisfaction[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Satisfaction.Employee"); // slightly different name
            EffortExpense[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense");
            EffortEngagement[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement");
            EffortGroupCoordination[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Coordination.Group"); // NOTE: Slightly different naming (due to length restrictions)!
            EffortGlobalCoordination[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Coordination.Global"); // NOTE: Slightly different naming
            EffortPlatformCoordination[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Coordination.Platform"); // NOTE: Slightly different naming
            ScoreBenefits[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits");
            ScoreEfforts[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts");
            ScoreComposite[0, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite");
            ImportanceCompliance[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.Average");
            ImportanceCPE[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.Average");
            ImportanceCost[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.Average");
            ImportanceCompete[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.Average");
            ImportanceConstraints[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.Average");
            ImportanceComplexity[0, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.Average");

            BenefitComplianceRisk[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.GSDS");
            BenefitEfficiency[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.GSDS");
            BenefitEffectiveness[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.GSDS");
            BenefitScale[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.GSDS");
            BenefitReach[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.GSDS");
            BenefitCustomerSatisfaction[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.GSDS");
            BenefitEmployeeSatisfaction[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.GSDS");
            EffortExpense[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.GSDS");
            EffortEngagement[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.GSDS");
            EffortGroupCoordination[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.GSDS");
            EffortGlobalCoordination[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.GSDS");
            EffortPlatformCoordination[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.GSDS");
            ScoreBenefits[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.GSDS");
            ScoreEfforts[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.GSDS");
            ScoreComposite[1, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.GSDS");
            ImportanceCompliance[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.GSDS");
            ImportanceCPE[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.GSDS");
            ImportanceCost[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.GSDS");
            ImportanceCompete[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.GSDS");
            ImportanceConstraints[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.GSDS");
            ImportanceComplexity[1, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.GSDS");

            BenefitComplianceRisk[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.AdSupport");
            BenefitEfficiency[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.AdSupport");
            BenefitEffectiveness[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.AdSupport");
            BenefitScale[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.AdSupport");
            BenefitReach[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.AdSupport");
            BenefitCustomerSatisfaction[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.AdSupport");
            BenefitEmployeeSatisfaction[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.AdSupport");
            EffortExpense[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.AdSupport");
            EffortEngagement[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.AdSupport");
            EffortGroupCoordination[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.AdSupport");
            EffortGlobalCoordination[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.AdSupport");
            EffortPlatformCoordination[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.AdSupport");
            ScoreBenefits[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.AdSupport");
            ScoreEfforts[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.AdSupport");
            ScoreComposite[2, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.AdSupport");
            ImportanceCompliance[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.AdSupport");
            ImportanceCPE[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.AdSupport");
            ImportanceCost[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.AdSupport");
            ImportanceCompete[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.AdSupport");
            ImportanceConstraints[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.AdSupport");
            ImportanceComplexity[2, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.AdSupport");

            BenefitComplianceRisk[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.CSS");
            BenefitEfficiency[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.CSS");
            BenefitEffectiveness[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.CSS");
            BenefitScale[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.CSS");
            BenefitReach[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.CSS");
            BenefitCustomerSatisfaction[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.CSS");
            BenefitEmployeeSatisfaction[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.CSS");
            EffortExpense[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.CSS");
            EffortEngagement[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.CSS");
            EffortGroupCoordination[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.CSS");
            EffortGlobalCoordination[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.CSS");
            EffortPlatformCoordination[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.CSS");
            ScoreBenefits[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.CSS");
            ScoreEfforts[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.CSS");
            ScoreComposite[3, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.CSS");
            ImportanceCompliance[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.CSS");
            ImportanceCPE[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.CSS");
            ImportanceCost[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.CSS");
            ImportanceCompete[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.CSS");
            ImportanceConstraints[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.CSS");
            ImportanceComplexity[3, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.Css");

            BenefitComplianceRisk[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.Finance");
            BenefitEfficiency[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.Finance");
            BenefitEffectiveness[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.Finance");
            BenefitScale[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.Finance");
            BenefitReach[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.Finance");
            BenefitCustomerSatisfaction[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.Finance");
            BenefitEmployeeSatisfaction[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.Finance");
            EffortExpense[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.Finance");
            EffortEngagement[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.Finance");
            EffortGroupCoordination[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.Finance");
            EffortGlobalCoordination[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.Finance");
            EffortPlatformCoordination[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.Finance");
            ScoreBenefits[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.Finance");
            ScoreEfforts[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.Finance");
            ScoreComposite[4, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.Finance");
            ImportanceCompliance[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.Finance");
            ImportanceCPE[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.Finance");
            ImportanceCost[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.Finance");
            ImportanceCompete[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.Finance");
            ImportanceConstraints[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.Finance");
            ImportanceComplexity[4, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.Finance");

            BenefitComplianceRisk[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.WOCs");
            BenefitEfficiency[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.WOCs");
            BenefitEffectiveness[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.WOCs");
            BenefitScale[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.WOCs");
            BenefitReach[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.WOCs");
            BenefitCustomerSatisfaction[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.WOCs");
            BenefitEmployeeSatisfaction[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.WOCs");
            EffortExpense[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.WOCs");
            EffortEngagement[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.WOCs");
            EffortGroupCoordination[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.WOCs");
            EffortGlobalCoordination[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.WOCs");
            EffortPlatformCoordination[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.WOCs");
            ScoreBenefits[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.WOCs");
            ScoreEfforts[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.WOCs");
            ScoreComposite[5, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.WOCs");
            ImportanceCompliance[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.WOCs");
            ImportanceCPE[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.WOCs");
            ImportanceCost[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.WOCs");
            ImportanceCompete[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.WOCs");
            ImportanceConstraints[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.WOCs");
            ImportanceComplexity[5, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.WOCs");

            // AdOps is the most important Team here, because they are the ones who 'own' the form
            // and its data.
            BenefitComplianceRisk[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.AdOps");
            BenefitEfficiency[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.AdOps");
            BenefitEffectiveness[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.AdOps");
            BenefitScale[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.AdOps");
            BenefitReach[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.AdOps");
            BenefitCustomerSatisfaction[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.AdOps");
            BenefitEmployeeSatisfaction[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.AdOps");
            EffortExpense[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.AdOps");
            EffortEngagement[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.AdOps");
            EffortGroupCoordination[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.AdOps");
            EffortGlobalCoordination[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.AdOps");
            EffortPlatformCoordination[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.AdOps");
            ScoreBenefits[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.AdOps");
            ScoreEfforts[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.AdOps");
            ScoreComposite[6, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.AdOps");
            ImportanceCompliance[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.AdOps");
            ImportanceCPE[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.AdOps");
            ImportanceCost[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.AdOps");
            ImportanceCompete[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.AdOps");
            ImportanceConstraints[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.AdOps");
            ImportanceComplexity[6, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.AdOps");

            BenefitComplianceRisk[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.TQ");
            BenefitEfficiency[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.TQ");
            BenefitEffectiveness[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.TQ");
            BenefitScale[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.TQ");
            BenefitReach[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.TQ");
            BenefitCustomerSatisfaction[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.TQ");
            BenefitEmployeeSatisfaction[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.TQ");
            EffortExpense[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.TQ");
            EffortEngagement[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.TQ");
            EffortGroupCoordination[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.TQ");
            EffortGlobalCoordination[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.TQ");
            EffortPlatformCoordination[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.TQ");
            ScoreBenefits[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.TQ");
            ScoreEfforts[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.TQ");
            ScoreComposite[7, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.TQ");
            ImportanceCompliance[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.TQ");
            ImportanceCPE[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.TQ");
            ImportanceCost[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.TQ");
            ImportanceCompete[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.TQ");
            ImportanceConstraints[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.TQ");
            ImportanceComplexity[7, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.TQ");

            BenefitComplianceRisk[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.ComplianceRisk.Accounting");
            BenefitEfficiency[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Efficiency.Accounting");
            BenefitEffectiveness[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Effectiveness.Accounting");
            BenefitScale[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Scale.Accounting");
            BenefitReach[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Reach.Accounting");
            BenefitCustomerSatisfaction[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.CustomerSat.Accounting");
            BenefitEmployeeSatisfaction[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.EmployeeSat.Accounting");
            EffortExpense[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Expense.Accounting");
            EffortEngagement[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Engagement.Accounting");
            EffortGroupCoordination[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Group.Accounting");
            EffortGlobalCoordination[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Global.Accounting");
            EffortPlatformCoordination[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Dimension.Platform.Accounting");
            ScoreBenefits[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Benefits.Accounting");
            ScoreEfforts[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Efforts.Accounting");
            ScoreComposite[8, 0] = wi.GetDecimalValue("Microsoft.Bios.WorkItem.StackRank.Total.Composite.Accounting");
            ImportanceCompliance[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Constraints.Accounting");
            ImportanceCPE[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compliance.Accounting");
            ImportanceCost[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.CPE.Accounting");
            ImportanceCompete[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Cost.Accounting");
            ImportanceConstraints[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Compete.Accounting");
            ImportanceComplexity[8, 0] = wi.GetDecimalValue("Microsoft.Operations.Project.Impact.Complexity.Accounting");

            TotalTeamSubmissions[0] = wi.GetFieldValue<int>("Microsoft.Bios.WorkItem.StackRank.Total.Submissions"); // loaded values

            // Other new items (requested November 2014)
            Priority[0] = wi.GetFieldValue<int>("Microsoft.VSTS.Common.Priority");
            EstimatedHours[0] = wi.GetFieldValue<double>("Microsoft.VSTS.Common.Priority");
            EffortLevelAdOps[0] = (wi.Fields["Microsoft.Bios.WorkItem.StackRank.Total.Efforts.AdOps.Level"].Value != null) ? wi.Fields["Microsoft.Bios.WorkItem.StackRank.Total.Efforts.AdOps.Level"].Value.ToString() : string.Empty;

            // The calculated values for 'TotalTeamSubmissions' is determined from the data loaded
            // for the big decimal arrays.
            // Rules: a total value of > 0 indicates at least one submission. Submission of ALL
            // ZEROES are not counted, since we might assume that is the same as being a completely
            // blank submission. (also because it is difficult to convert all the decimal stuff above
            // to work with NULL values. If you check out 'LoadFromField' you will note we are
            // treating 'null' as equivalent to Zero (simply because it makes it easier to work with
            // the numbers and it's more accurate too). We have to check all the items which make up
            // the score, we can't trust 'Composite Score' since it may have some other calculation
            // later on. There are 12 values we have to check, x 8 teams each. we can ignore Team
            // Zero, because that is used for averages. We have to check from the LOADED values.

            int teamCount = 0;

            for (int i = 1; i < 9; i++)
            {
                if (( // add up all these decimal values and ...
                BenefitComplianceRisk[i, 0] +
                BenefitEfficiency[i, 0] +
                BenefitEffectiveness[i, 0] +
                BenefitScale[i, 0] +
                BenefitReach[i, 0] +
                BenefitCustomerSatisfaction[i, 0] +
                BenefitEmployeeSatisfaction[i, 0] +
                EffortExpense[i, 0] +
                EffortEngagement[i, 0] +
                EffortGroupCoordination[i, 0] +
                EffortGlobalCoordination[i, 0] +
                EffortPlatformCoordination[i, 0]
                ) > 0) teamCount++; // if greater than zero then increment the team count.
            }

            TotalTeamSubmissions[1] = teamCount; // calculated values

            ///////////// find the averages first, for each.
            // According to Jason Parish (July 2014), it is adequate to use divided by 8 teams to get
            // the average. It does not matter whether a value has been submitted or not. Values
            // should be to one decimal place, for simplicity.

            int decPlaces = 1;

            BenefitComplianceRisk[0, 1] = decimal.Round(AverageFrom(BenefitComplianceRisk, 0, teamCount), decPlaces);
            BenefitEfficiency[0, 1] = decimal.Round(AverageFrom(BenefitEfficiency, 0, teamCount), decPlaces);
            BenefitEffectiveness[0, 1] = decimal.Round(AverageFrom(BenefitEffectiveness, 0, teamCount), decPlaces);
            BenefitScale[0, 1] = decimal.Round(AverageFrom(BenefitScale, 0, teamCount), decPlaces);
            BenefitReach[0, 1] = decimal.Round(AverageFrom(BenefitReach, 0, teamCount), decPlaces);
            BenefitCustomerSatisfaction[0, 1] = decimal.Round(AverageFrom(BenefitCustomerSatisfaction, 0, teamCount), decPlaces);
            BenefitEmployeeSatisfaction[0, 1] = decimal.Round(AverageFrom(BenefitEmployeeSatisfaction, 0, teamCount), decPlaces);
            EffortExpense[0, 1] = decimal.Round(AverageFrom(EffortExpense, 0, teamCount), decPlaces);
            EffortEngagement[0, 1] = decimal.Round(AverageFrom(EffortEngagement, 0, teamCount), decPlaces);
            EffortGroupCoordination[0, 1] = decimal.Round(AverageFrom(EffortGroupCoordination, 0, teamCount), decPlaces);
            EffortGlobalCoordination[0, 1] = decimal.Round(AverageFrom(EffortGlobalCoordination, 0, teamCount), decPlaces);
            EffortPlatformCoordination[0, 1] = decimal.Round(AverageFrom(EffortPlatformCoordination, 0, teamCount), decPlaces);

            // Task 12156 .. add scoring for 'FMX Total Score', which is a rating for their external
            // ranking. Complex formulae is supplied by Jason Parrish (Revel Consulting) <v-jaspar@microsoft.com>

            FMXUsage[0] = (wi.Fields["Microsoft.Bios.WorkItem.ID.ExternalReference.IsFMX"].Value != null) ? wi.Fields["Microsoft.Bios.WorkItem.ID.ExternalReference.IsFMX"].Value.ToString() : string.Empty;

            if (FMXUsage[0] == "Yes") // then we care about the score
            {
                FMXTotalScore[0] = wi.GetFieldValue<decimal>("Microsoft.Bios.WorkItem.StackRank.Dimension.FMX");
                FMXTotalScore[1] = (BenefitComplianceRisk[0, 1] + BenefitCustomerSatisfaction[0, 1] + BenefitCustomerSatisfaction[0, 1] + 1 + ((BenefitEfficiency[0, 1] + BenefitEffectiveness[0, 1]) / 2)); // NOTE: Yes, customer satisfaction is intentionally added twice.
            }

            // Calculate the Benefit Scores for each TEAM SCORING SUBMISSION, based on LOADED values
            // Note that some limit of reading error is introduced because the original averages are
            // also rounded, but this is permissible.

            for (int i = 0; i < 9; i++)
            {
                ScoreBenefits[i, 1] = decimal.Round(BenefitComplianceRisk[i, 0] + BenefitEfficiency[i, 0] + BenefitEffectiveness[i, 0] + BenefitScale[i, 0] + BenefitReach[i, 0] + BenefitCustomerSatisfaction[i, 0] + BenefitEmployeeSatisfaction[i, 0], decPlaces);
                ScoreEfforts[i, 1] = decimal.Round(EffortExpense[i, 0] + EffortEngagement[i, 0] + EffortGroupCoordination[i, 0] + EffortGlobalCoordination[i, 0] + EffortPlatformCoordination[i, 0], decPlaces);
                ScoreComposite[i, 1] = decimal.Round(ScoreBenefits[i, 1] - ScoreEfforts[i, 1], decPlaces); // CALCULATED derived from other CALCULATED values

                // Also perform a calculation for the decimal values of Project priority, etc. These
                // CALCULATED values are likewise determined from the LOADED values of all the
                // Benefit/Effort metrics supplied.

                // From: Jason Parrish (Revel Consulting)
                // Sent: Saturday, 21 June 2014 6:54 AM
                // To: Warren James (Adecco)
                // Cc: Harriet Smith
                // Subject: ABO Planning Tool - 6 C's Calculations

                // Hi Warren,

                // Harriet and I came up with the following formulas to calculate the 6 C’s in the tool.
                // 1.       Compliance: Compliance score
                // 2.       CPE: Customer Satisfaction score
                // 3.       Cost: ((Cost+Engagement+Group Coordination+Global Coordination+Platform Coordination)-(Efficiency+Effectiveness))/7
                // 4.       Compete: (Scalability+Reach+Effectiveness)/3
                // 5.       Complexity: (Engagement+Group Coordination+Global Coordination+Platform Coordination)/4
                // 6.       Constraints: (Cost+Engagement+Group Coordination+Global
                //          Coordination+Platform Coordination)/5

                // Following each of these calculations we would then need to apply the H/M/L value
                // using the predetermined range (0,1 = Low; 2,3 = Medium; 4,5 = High).

                ImportanceCompliance[i, 1] = BenefitComplianceRisk[i, 0];
                ImportanceCPE[i, 1] = BenefitCustomerSatisfaction[i, 0];
                ImportanceCost[i, 1] = decimal.Round(((EffortExpense[i, 0] + EffortGroupCoordination[i, 0] + EffortGlobalCoordination[i, 0] + EffortPlatformCoordination[i, 0]) - (BenefitEfficiency[i, 0] + BenefitEffectiveness[i, 0])) / 7, decPlaces);
                ImportanceCompete[i, 1] = decimal.Round((BenefitScale[i, 0] + BenefitReach[i, 0] + BenefitEffectiveness[i, 0]) / 3, decPlaces);
                ImportanceConstraints[i, 1] = decimal.Round((EffortExpense[i, 0] + EffortGroupCoordination[i, 0] + EffortGlobalCoordination[i, 0] + EffortPlatformCoordination[i, 0]) / 5, decPlaces);
                ImportanceComplexity[i, 1] = decimal.Round((EffortEngagement[i, 0] + EffortGroupCoordination[i, 0] + EffortPlatformCoordination[i, 0]) / 4, decPlaces);
            }

            // Finally, we can now get the overall average score. Note, these averages should be from
            // the CALCULATED figures, not the LOADED figures.

            ScoreBenefits[0, 1] = decimal.Round(AverageFrom(ScoreBenefits, 1, teamCount), decPlaces);
            ScoreEfforts[0, 1] = decimal.Round(AverageFrom(ScoreEfforts, 1, teamCount), decPlaces);
            ScoreComposite[0, 1] = decimal.Round(AverageFrom(ScoreComposite, 1, teamCount), decPlaces);

            ImportanceCompliance[0, 1] = decimal.Round(AverageFrom(ImportanceCompliance, 1, teamCount), decPlaces);
            ImportanceCPE[0, 1] = decimal.Round(AverageFrom(ImportanceCPE, 1, teamCount), decPlaces);
            ImportanceCost[0, 1] = decimal.Round(AverageFrom(ImportanceCost, 1, teamCount), decPlaces);
            ImportanceCompete[0, 1] = decimal.Round(AverageFrom(ImportanceCompete, 1, teamCount), decPlaces);
            ImportanceConstraints[0, 1] = decimal.Round(AverageFrom(ImportanceConstraints, 1, teamCount), decPlaces);
            ImportanceComplexity[0, 1] = decimal.Round(AverageFrom(ImportanceComplexity, 1, teamCount), decPlaces);

            // And finally, haven't forgotten about them .. Make the calculation on the main impact
            // (importance) scoring. Please note that the value is based on ADOPS Scores, rather than
            // AVERAGE, because the averages will always be quite low based on their 'divide by
            // eight' methodology. The importance is all about ADOPS.

            ImportanceComplianceText[1] = TShirtSizeBasedOn(ImportanceCompliance[0, 1]);
            ImportanceCPEText[1] = TShirtSizeBasedOn(ImportanceCPE[0, 1]);
            ImportanceCostText[1] = TShirtSizeBasedOn(ImportanceCost[0, 1]);
            ImportanceCompeteText[1] = TShirtSizeBasedOn(ImportanceCompete[0, 1]);
            ImportanceConstraintsText[1] = TShirtSizeBasedOn(ImportanceConstraints[0, 1]);
            ImportanceComplexityText[1] = TShirtSizeBasedOn(ImportanceComplexity[0, 1]);

            // These calculations added Nov 2014, as per Task 14928.
            Priority[1] = (ScoreComposite[0, 1] > 7) ? 1 : (ScoreComposite[0, 1] > 3) ? 2 : 3; // according to their logic (see workitem)
            EffortLevelAdOps[1] = (ScoreEfforts[6, 0] >= 21) ? "Extra High" : (ScoreEfforts[6, 0] > 15) ? "High" : (ScoreEfforts[6, 0] > 8) ? "Medium" : "Low";
            EstimatedHours[1] = (ScoreEfforts[6, 0] >= 21) ? 1490 : (ScoreEfforts[6, 0] > 15) ? 1070 : (ScoreEfforts[6, 0] > 8) ? 555 : 208; // is based on above

            Calculate(1); // now that we have values
        }

        /// <summary>
        /// Compares a 'loaded' value with the same 'calculated' value.
        /// </summary>
        private static bool LoadedEqualsCalculated<T>(T[,] array)
        {
            EqualityComparer<T> comparer = EqualityComparer<T>.Default;
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if (!comparer.Equals(array[i, 0], array[i, 1])) return false;
            }
            return true;
        }

        private static bool LoadedEqualsCalculated<T>(T[] one_dimensional_array)
        {
            EqualityComparer<T> comparer = EqualityComparer<T>.Default;
            if (!comparer.Equals(one_dimensional_array[0], one_dimensional_array[1])) return false;
            return true;
        }

        /// <summary>
        /// NOTE: No provision for negative numbers (0,1 = Low; 2,3 = Medium; 4,5 = High).
        /// </summary>
        private static string TShirtSizeBasedOn(decimal numericScore)
        {
            string size = string.Empty;

            if (numericScore == 0)
                size = "None";
            else if (numericScore <= 1)
                size = "Low";
            else if (numericScore <= 3)
                size = "Medium";
            else if (numericScore > 3) // includes 4,5 but also allows for higher numbers due to the way Jason's scoring works
                size = "High"; // highest

            return size;
        }

        /// <summary>
        /// Use our own Average method (rather than LINQ) because we have a specialized requirement
        /// </summary>
        private decimal AverageFrom(decimal[,] array, int fromPosition, int sampleSize)
        {
            decimal total = 0;

            for (int index = 1; index < array.GetLength(0); index++)
            {
                total += array[index, fromPosition];
            }

            // return (decimal)total / (array.GetLength(0) - 1); // don't include original averages
            // 'column zero'
            if (sampleSize == 0)
            {
                return 0;
            }
            else
            {
                return total / sampleSize; // new in November 2014, averages come from Team Count
            }
        }
    }
}