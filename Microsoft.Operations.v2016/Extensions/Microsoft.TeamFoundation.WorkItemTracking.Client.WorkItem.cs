using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections;

namespace Microsoft.Operations
{
    public static partial class TeamFoundationServerHelper
    {
        /// <summary>
        /// Shallow clone, copies across attributes (v. basic)
        /// </summary>
        /// <param name="wi">The workitem you wish to make a copy of</param>
        /// <remarks>Probably a better way of doing this (later)</remarks>
        public static WorkItemClone Clone(this WorkItem wi)
        {
            WorkItemClone swi = new WorkItemClone();
            swi.Id = wi.Id;
            swi.Title = wi.Title;
            swi.Type = wi.Type;
            swi.ChangedDate = wi.ChangedDate;
            swi.ChangedBy = wi.ChangedBy;
            swi.Uri = wi.Uri;
            swi.State = wi.State;
            return swi;
        }

        public static decimal GetDecimalValue(this WorkItem wi, string referenceName)
        {
            if (wi.Fields[referenceName].Value == null)
            {
                return Convert.ToDecimal(0);
            }
            else
            {
                return Convert.ToDecimal(wi.Fields[referenceName].Value);
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="wi"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
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

        public static T GetFieldValue<T>(this WorkItem wi, string referenceName)
        {
            try
            {
                return (wi.Fields[referenceName].Value != null) ? (T)wi.Fields[referenceName].Value : default(T);
            }
            catch (Exception ex)
            {
                return default(T);
            }
        }

        /// <summary>
        /// Shorthand check for the 'Changed Date' value, compared to current date/time.
        /// </summary>
        public static bool LastChangedAgeInMinutesIsAtLeastMoreThan(this WorkItem wi, int numberOfMinutes)
        {
            return (wi.ChangedDate.AddMinutes(numberOfMinutes) < DateTime.Now) ? true : false;
        }

        /// <summary>
        /// OUR CUSTOM VALIDATION routine, useful for Windows Services because it automatically sends
        /// (us) a notification if there are any errors. This helps us keep on top of any
        /// template/rules violations which are otherwise difficult in a headless environment!
        ///
        /// The routine itself performs the following:
        /// 1. Determines if the Work Item is 'dirty' (i.e. changes have actually been made) if not,
        ///    returns false right away.
        /// 2. Verifies there are zero validation errors from the inbuilt 'Validate()' method.
        /// 3. Emails any error messages with some information
        ///
        /// If errors are found, an email will automatically go to the DevOps group nominated, using
        /// a specially formatted email with all the details of the process.
        ///
        /// TODO: this routine would be more useful if it had the option of returning the 'error set'
        /// </summary>
        public static bool ValidatesOkay(this WorkItem wi, AzureLogger logging = null)
        {
            if (wi.IsDirty)
            {
                // we have a potentially 'saveable' workitem .. now check for errors:

                ArrayList result = wi.Validate();

                if (result.Count == 0)
                {
                    return true;
                }
                else
                {
                    // Errors exist, persist to log ... if it's been given.
                    if (logging != null)
                    {
                        foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result)
                        {
                            logging.WriteLine(string.Format("ID:{0}, Field:[{1}] {2}", wi.Id, info.Name, info.Status));
                        }
                    }

                    // Write a queue message to show that we have an error

                    return false;
                }
            }
            else
            {
                // A work item which hasn't been changed is not a candidate for saving. Note that
                // although there are typically no validation errors at this point, it is still
                // possible for validation errors to exist (e.g. if the template has changed and the
                // new rules have a conflict) but that scenario is not in scope for this function.
                return false;
            }
        }
    }
}