using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;

namespace Microsoft.Operations.CSP.RegSys
{
    [Serializable]
    public class Invoice : WorkItemSystemFields
    {
        public string AssignedTo;

        public string BillAmount;

        public string BillFrequency;

        public string CMATDomain;

        public string Currency;

        public string DesireBillDate;

        public string FirstBillDate;

        public string HQMPNID;

        public string InvoiceDeliveryMethod;

        public string IsSAP;

        public List<InvoiceDetail> LineItems;

        public string PartnerAddressLine1;

        public string PartnerAddressLine2;

        public string PartnerBillingAddressLine1;

        public string PartnerBillingAddressLine2;

        public string PartnerBillingCity;

        public string PartnerBillingContactName;

        public string PartnerBillingCountry;

        public string PartnerBillingEmail;

        public string PartnerBillingState;

        public string PartnerBillingZip;

        public string PartnerCity;

        public string PartnerContactEmail;

        public string PartnerContactName;

        public string PartnerCountry;

        public string PartnerOrg;

        public string PartnerShipAddressLine1;

        public string PartnerShipAddressLine2;

        public string PartnerShipCity;

        public string PartnerShipContactName;

        public string PartnerShipCountry;

        public string PartnerShipEmail;

        public string PartnerShipState;

        public string PartnerShipZip;

        public string PartnerState;

        public string PartnerZip;

        public string PaymentTerms;

        public string ServiceEndDate;

        public string ServiceStartDate;

        /// <summary>
        /// The original SKU information (as text) which came from the original submission.
        /// </summary>
        public string SKUs_OriginalInformation;

        public string Tax_VATID;

        public string TaxStatus;

        public Invoice()
        {
            LineItems = new List<InvoiceDetail>();
        }

        public Invoice(WorkItem wi) : base()
        {
            wi.PartialOpen();
        }
    }
}