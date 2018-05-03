using System;

/// <summary>
/// Contains just the system fields common to a TFS workitem.
/// </summary>
public class WorkItemSystemFields
{
    public string ChangedBy;

    /// <summary>
    /// Always has a value &gt; when the workitem was last saved
    /// </summary>
    public DateTime ChangedDate;

    /// <summary>
    /// History value (text) which would get written
    /// </summary>
    public string History;

    public int Id;

    /// <summary>
    /// System.State, which is the primary driver of all workflow activity
    /// </summary>
    public string State;

    public string Title;

    /// <summary>
    /// Location of the Web URL for this particular invoice record i.e. specifically via Team System
    /// Web Access
    /// </summary>
    public Uri Uri;

    public string WorkItemTypeName;
}