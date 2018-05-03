using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;

public static partial class TeamFoundationServerHelper
{
    public static string GetFieldValue(this WorkItem wi, string referenceName)
    {
        try
        {
            return (wi.Fields[referenceName].Value != null) ? wi.Fields[referenceName].Value.ToString() : string.Empty;
        }
        catch (Exception)
        {
            // if fields value not found, the error TF26027 will be catch.
            // Exp: TF26027: A field definition System.Links.LinkType in the work item type
            //      definition file does not exist. Add a definition for this field or remove the
            //      reference to the field and try again. Error handling: Return empty instead
            return string.Empty;
        }
    }

    /// <summary>
    /// Shorthand check for the 'Changed Date' value, compared to current date/time.
    /// </summary>
    public static bool LastChangedAgeInMinutes(this WorkItem wi, int numberOfMinutes)
    {
        return (wi.ChangedDate.AddMinutes(numberOfMinutes) < DateTime.Now) ? true : false;
    }
}