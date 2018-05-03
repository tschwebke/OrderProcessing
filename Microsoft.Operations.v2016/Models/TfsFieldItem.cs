using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Microsoft.Operations.TeamFoundationServer
{
    /// <summary>
    /// A shallow copy of the FieldDefinition object. Use for cataloging.
    /// </summary>
    public class FieldItem
    {
        public FieldItem()
        {
        }

        public FieldItem(FieldDefinition field)
        {
            ID = field.Id;
            Name = field.Name;
            FieldType = field.FieldType;
            IsCoreField = field.IsCoreField;
            ReferenceName = field.ReferenceName;
            ReportingAttributes = field.ReportingAttributes;//TFS 2010
            FieldUsages = field.Usage; //TFS 2010
            IsEditable = field.IsEditable;
            AllowedValuesCollection = field.AllowedValues;
        }

        public AllowedValuesCollection AllowedValuesCollection { get; set; }
        public FieldType FieldType { get; set; }
        public FieldUsages FieldUsages { get; set; }
        public string HelpText { get; set; }
        public int ID { get; set; }
        public bool IsCoreField { get; set; }
        public bool IsEditable { get; set; }
        public string Name { get; set; }
        public string ReferenceName { get; set; }
        public ReportingAttributes ReportingAttributes { get; set; }

        public override string ToString()
        {
            return string.Format("{0} ({1})", this.Name, this.ReferenceName);
        }
    }
}