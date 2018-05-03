using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;

namespace Microsoft.Operations
{
    /// <summary>
    /// The normal TFS API WorkItem object is sealed. This class allows you to extend it, but
    /// requires copying as a shallow clone.
    /// </summary>
    public class WorkItemClone
    {
        private string _changed_by;
        private DateTime _date_changed;
        private DateTime? _date_finish = null;
        private DateTime? _date_start = null;
        private List<int> _error_codes;
        private int _id;

        // custom fields
        private int _severity;

        private string _state;
        private Uri _tfs_web_location_this_workitem;
        private string _title;
        private WorkItemType _type;

        public string ChangedBy
        {
            get { return _changed_by; }
            set { _changed_by = value; }
        }

        public DateTime ChangedDate
        {
            get { return _date_changed; }
            set { _date_changed = value; }
        }

        public List<int> Errors
        {
            get { return _error_codes; }
            set { _error_codes = value; }
        }

        public DateTime? FinishDate
        {
            get { return _date_finish; }
            set { _date_finish = value; }
        }

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int NonComplianceSeverity
        {
            get { return _severity; }
            set { _severity = value; }
        }

        public DateTime? StartDate
        {
            get { return _date_start; }
            set { _date_start = value; }
        }

        public string State
        {
            get { return _state; }
            set { _state = value; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        public WorkItemType Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public Uri Uri
        {
            get { return _tfs_web_location_this_workitem; }
            set { _tfs_web_location_this_workitem = value; }
        }
    }
}