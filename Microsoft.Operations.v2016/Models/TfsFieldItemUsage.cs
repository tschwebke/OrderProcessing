using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Operations.TeamFoundationServer
{
    /// <summary>
    /// Field usage with a whole lot more information attached (used for reporting)
    /// </summary>
    public class FieldItemUsage : FieldItem
    {
        private Dictionary<int, int> _density = new Dictionary<int, int>();
        private Dictionary<string, int> _states = new Dictionary<string, int>();

        public FieldItemUsage()
        {
            DateFirstUsage = DateTime.Now; // put a benchmark, on which the date can be moved down
            _density.Add(2, 0);
            _density.Add(4, 0);
            _density.Add(8, 0);
            _density.Add(16, 0);
            _density.Add(32, 0);
            _density.Add(64, 0);
            _density.Add(128, 0);
            _density.Add(256, 0);
            _density.Add(512, 0);
            _density.Add(1024, 0);
        }

        public FieldItemUsage(FieldDefinition field)
        {
            ID = field.Id;
            Name = field.Name;
            FieldType = field.FieldType;
            ReferenceName = field.ReferenceName;
            ReportingAttributes = field.ReportingAttributes;//TFS 2010
            FieldUsages = field.Usage; //TFS 2010
            IsEditable = field.IsEditable;
            IsCoreField = field.IsCoreField;
            AllowedValuesCollection = field.AllowedValues;
            DateFirstUsage = DateTime.Now; // put a benchmark, on which the date can be moved down
            _density.Add(2, 0);
            _density.Add(4, 0);
            _density.Add(8, 0);
            _density.Add(16, 0);
            _density.Add(32, 0);
            _density.Add(64, 0);
            _density.Add(128, 0);
            _density.Add(256, 0);
            _density.Add(512, 0);
            _density.Add(1024, 0);
            _density.Add(2048, 0);
        }

        public DateTime DateFirstUsage { get; set; }

        public DateTime DateLastUsage { get; set; }

        public Dictionary<int, int> Density
        {
            get { return _density; }
            set { _density = value; }
        }

        public int HighestWorkItemID { get; set; }

        public bool IsFieldInABO_Layout { get; set; }

        public bool IsFieldInSBO_Layout { get; set; }

        /// <summary>
        /// Simple flag to indicate whether the field is contained in the layout or not.
        /// </summary>
        public bool IsInLayout { get; set; }

        public bool IsInTemplate { get; set; }

        /// <summary>
        /// Simple flag to indicate whether the field is uses as part of the workflow in a certain template
        /// </summary>
        public bool IsInWorkflow { get; set; }

        /// <summary>
        /// Number which indicates the order of appearance of a field in the <![CDATA[<LAYOUT>]]>
        /// section of a work item template.
        /// </summary>
        public int OrderOfAppearanceInLayout
        {
            get;
            set;
        }

        public Dictionary<string, int> StateUsage
        {
            get { return _states; }
            set { _states = value; }
        }

        public string StateUsageSummary
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach (string s in StateUsage.Keys)
                {
                    sb.Append(string.Format("{0} {1}, ", s, StateUsage[s]));
                }
                return sb.ToString().Trim().TrimEnd(',');
            }
        }

        public int UsageTally { get; set; }

        public decimal UsagePercentage(int totalSamples)
        {
            if (totalSamples == 0)
            {
                return 0;
            }
            else
            {
                return (Convert.ToDecimal(UsageTally) / Convert.ToDecimal(totalSamples));
            }
        }
    }
}