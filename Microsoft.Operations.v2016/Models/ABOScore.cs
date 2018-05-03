using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Data.Services.Common;

/// <summary>
/// A scoring entity (for a specific ABO sub-group), used for transferring values to and from the
/// Azure scoring functionality.
/// </summary>
[DataServiceKey("PartitionKey", "RowKey")]
public class ABOScore : TableEntity
{
    /// <summary>
    /// New ABO S
    /// </summary>
    public ABOScore() { }

    public ABOScore(string partitionKey, string rowKey)
    {
        PartitionKey = partitionKey;
        RowKey = rowKey;
    }

    public string Author { get; set; }
    public string BenefitComplianceRisk { get; set; }
    public string BenefitCustomerSatisfaction { get; set; }
    public string BenefitEffectiveness { get; set; }

    // These values below are actually all decimals, but we have to cater for the null value in TFS
    // so it's easier in terms of cloud storage + website updating, to treat them as strings.
    public string BenefitEfficiency { get; set; }

    public string BenefitEmployeeSatisfaction { get; set; }
    public string BenefitReach { get; set; }
    public string BenefitScalability { get; set; }
    public string BusinessCaseJustification { get; set; }
    public DateTime ChangedDate { get; set; }
    public string Comments { get; set; }

    // this may possibly be later, a common field (it should be)
    public string EffortCost { get; set; }

    public string EffortEngagement { get; set; }
    public string EffortGlobalCoordination { get; set; }
    public string EffortGroupCoordination { get; set; }
    public string EffortPlatformCoordination { get; set; }
    public int Id { get; set; }
    public string InformationOnMandatoryForComplianceLegalRegulatory { get; set; }
    public string MandatoryForComplianceLegalRegulatory { get; set; }
    public string OperationsBenefits { get; set; }
    public string ProblemStatement { get; set; }

    // added 2015 to support ABO Scoring website added 2015 to support ABO Scoring website added 2015
    // to support ABO Scoring website
    public string Recommendation { get; set; }

    public string Requestor { get; set; }

    // added 2015 to support ABO Scoring website
    public string RevenueImpact { get; set; }

    // added 2015 to support ABO Scoring website
    public string RiskIfNotImplemented { get; set; }

    public string Title { get; set; }
    public string WebsiteChangedBy { get; set; }
    public bool WebsiteIsModified { get; set; }
}