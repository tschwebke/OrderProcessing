using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;

namespace Microsoft.Operations.CSP.RegSys
{
    [Serializable]
    public class InvoiceDetail : WorkItemSystemFields
    {
        public double Amount;

        public string BillDate;

        public string BillFrequency;

        public string Description;

        public string EndDate;

        public int Quantity;

        public string SKU;

        public string StartDate;

        public InvoiceDetail()
        {
        }

        public InvoiceDetail(WorkItem wi)
        {
            wi.PartialOpen();
        }
    }
}