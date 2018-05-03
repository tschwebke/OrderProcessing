using System;

/// <summary>
/// This item is used for fetching SharePoint submissions. The 'object' is used specifically because
/// the data types in SharePoint are nebulous.
/// </summary>
public class AdopsIssue
{
    public object AdditionalInformation;

    public object AnySpecificFeature;

    public object BusinessJustification;

    public object Category;

    public object ComplianceLegalRegulatoryOptions;

    public object DateModified;

    public object EngineeringTeamsImpacted;

    // multiselect
    public object InformationOnComplianceImpact;

    public object MandatoryForComplianceLegalRegulatory;

    public object NonEngineeringTeamsImpacted;

    public object OperationsBenefits;

    public object ProblemStatement;

    public object Recommendation;

    public object RevenueImpact;

    public object RiskImpactIfNotDone;

    // yes/no
    public object ScoreComplianceRisk;

    public object ScoreCost;

    public object ScoreCustomerSatisfaction;

    public object ScoreEffectiveness;

    public object ScoreEfficiency;

    public object ScoreEmployeeSatisfaction;

    public object ScoreEngagement;

    public object ScoreGlobalCoordination;

    public object ScoreGroupCoordination;

    public object ScorePlatformCoordination;

    public object ScoreReach;

    public object ScoreScalability;

    public object SubjectCategory1;

    public object SubjectCategory2;

    public object SubmissionDate;

    public object SubmissionTeam;

    public object SubmittedBy;

    public string SubmittedByEmailAddress;

    public string SubmittedByUserName;

    public object Title;

    public Uri UrlLink;

    public AdopsIssue()
    {
        Title = string.Empty;
    }
}