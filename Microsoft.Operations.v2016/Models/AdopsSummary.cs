using System.Collections.Generic;

/// <summary>
/// This item is used for fetching SharePoint submissions. The 'object' is used specifically because
/// the data types in SharePoint are nebulous.
/// </summary>
public class AdopsSummary : WorkItemSystemFields
{
    public object BusinessJustification;

    public List<string> IssuesRisks;

    public List<string> KeyAccomplishments;

    public object OperationsBenefits;

    public object ProblemStatement;

    public object Recommendation;

    public object RiskImpactIfNotDone;

    /// <summary>
    /// Also known as 'Color' or color Indicator. i.e. Green, Red, Yellow.
    /// </summary>
    public string Status;

    public AdopsSummary()
    {
        KeyAccomplishments = new List<string>();
    }
}