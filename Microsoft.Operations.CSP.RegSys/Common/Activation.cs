using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;

namespace Microsoft.Operations.CSP.RegSys
{
    [Serializable]
    public class Activation : WorkItemSystemFields
    {
        public string AssignedTo;

        public string CustomBillAmount;

        public string FormLanguage;

        public string HQ_Country;

        public string HQ_MPNID;

        public string PartnerBillingContact;

        public string PartnerBillingEmail;

        public string PartnerContactEmail;

        public string PartnerCurrentState;

        public string PartnerName;

        public string PartnerOrg;

        public string PartnerShipEmail;

        public string SAM;

        public string SAMEmail;

        public string SAMLead;

        public string SAMLeadEmail;

        public DateTime ServiceStartDate;

        public string SKUs;

        public DateTime TrialServiceEndDate;

        public DateTime TrialServiceStartDate;

        public string VATIDValue;

        public Activation()
        {
        }

        public Activation(WorkItem wi) : base()
        {
            wi.PartialOpen();
        }
    }
}