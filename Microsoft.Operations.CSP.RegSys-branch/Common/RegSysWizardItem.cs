using Microsoft.Operations;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

[Serializable]
public class RegSysWizardItem : WorkItemSystemFields
{
    /// <summary>
    /// Any information about why the item is not compliant. Any entries here should indicate the
    /// item is malformed. Items with errors will NOT be processed.
    /// </summary>
    public List<string> Errors;

    public List<string> Information;

    public string PartnerGUID;

    /// <summary>
    /// Carries information which might be low-level impact but is not serious enough to stop
    /// processing. This was introduced to warn of duplicate entries.
    /// </summary>
    public List<string> Warnings;

    public RegSysWizardItem()
    {
    }

    /// <summary>
    /// Attempts to load all the data for an item, from an Xml fragment. This is based on a
    /// pre-defined setup from the regsys information.
    /// </summary>
    public RegSysWizardItem(XmlElement xml) : base()
    {
        SKUs = string.Empty;
        PartnerOrganizationName = string.Empty;
        if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "chads")
        {
            DataOriginDetail = "https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=84dd6e58-e97a-4091-b632-bb0df8eb4a48&lcid=9";
        }
        else
        {
            DataOriginDetail = "https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=8d076345-5e8c-4e3f-9d7a-1fc3c9be69fd&lcid=9";
        }
        DataOrigin = "ASfP RegSys Form";
        //DataOriginDetail = "https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=8d076345-5e8c-4e3f-9d7a-1fc3c9be69fd&lcid=9";

        // ReportUrl = "http://co1msftolappa02/AuthCustomerResponses/public/index.html?WizId=8d076345-5e8c-4e3f-9d7a-1fc3c9be69fd";

        Errors = new List<string>();
        Warnings = new List<string>();
        Information = new List<string>();

        foreach (XmlElement pair in xml.FirstChild.ChildNodes)
        {
            string key = pair.FirstChild.InnerText;
            string value = pair.LastChild.InnerText;

            switch (key)
            {
                // VERY IMPORTANT!!!! The 'Created Date' value is relative to the requester.
                // Therefore, if the request originates from a location different to Redmond then the
                // date will have a different value. This impacts testing and processing because we
                // use the date as part of our virtual key.

                case "WizardID": WizardGUID = value; break;
                case "Lcid": LCID = value; break;
                case "IpAdress": IPAddress = value; break; // yes, they've spelt it incorrectly!
                // case "CustomerID": PartnerGUID = value; break; // the customer is actually the
                // partner ...
                case "IsAuthenticated": IsAuthenticated = value; break;
                case "CreatedDateTime":
                    // example is: 10/14/2015 5:19:19 PM
                    DateTime test;
                    bool successful = DateTime.TryParseExact(value.ToString(), @"M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out test); // on test server, australian.
                    if (successful) CreatedDateInRegSys = test.ToUniversalTime();
                    break;

                case "IsEmailVerified": IsEmailVerified = value; break;

                case "QID698043afcfc848d99adf30019088d7bc":
                case "QIDc7aa2740c3cf45db98deef24bb3f8838":
                case "QIDb6bef45f8bb74610a054e56923f0654b":
                case "QIDca96efe55b8c4b27bee10d50846e1efa":
                    // These 4 are all specifically-named SKU items. You'll have to add to this
                    // section if any other ones get created.
                    if (!string.IsNullOrEmpty(value))
                    {
                        SKUs += value + Environment.NewLine;
                    }
                    break;

                case "QID2b75a947864e41f3970f078dd099974a": SpecialInstructions = value; break;
                case "QID529fc10c5afe45e2ae815efcc029877d": MPNID = value; break;
                case "QID5f48d693c4d34c149aa0398c477c6e13": CustomerTenantDomain = value; break;
                case "QID123ab46f5e564a659705617853dc3436": PartnerContactName = value; break;
                case "QIDa20cad709cdf49f692f7227c117d59e2": PartnerContactEmail = value; break;
                case "QIDcc9d6f8bf1e647dfbce07b88daafa817": PartnerArea = value; break;
                case "QIDe05f0c99ad6b4ad1ab027f199f09ce06": BillingContactName = value; break;
                case "QID90a12ec1b9f94cb6886fd54c42eb7adc": BillingContactEmail = value; break;
                case "QIDea3cdc0d3de94fe39c638fe226fc8583":
                    AdditionalEmail = value; break;

                    // Rather than having two separate fields for Billing Information, we'll append
                    // the overflow to the BillingContactEmail field. Purpose retained, but for less cost.

                    if (!string.IsNullOrEmpty(value) && string.IsNullOrEmpty(BillingContactEmail)) BillingContactEmail = value; // use this instead (rare)
                    if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(BillingContactEmail)) BillingContactEmail = BillingContactEmail + "," + value; // append the values
                    break;

                case "QID88a590c2da254c39847af73da3be7b94": AgreementToTerms = value; break;
            }
        }

        // Now apply some cleanups to the raw form data which was collected.

        DataOrigin = "ASfP RegSys Form";

        // RegSys doesn't provide a unique key for the record (wonder why?) Our virtual key format
        // will cover this requirement. The format will only ever bomb if we have more than one
        // submission from the same partner *within the same second* That would be an unlikely event,
        // end user would need to submit the form twice in the same second (and we assume regsys
        // itself has mechanisms to prevent that kind of duplication).

        if (!string.IsNullOrEmpty(BillingContactEmail))
        {
            BillingContactEmail = BillingContactEmail.Replace(",", ";");
            BillingContactEmail = BillingContactEmail.Trim();
        }

        string unique_id = PartnerContactEmail.Substring(0, 1);
        PartnerContactEmail = PartnerContactEmail.Trim();
        if (PartnerContactEmail.IsValidEmailAddress())
        {
            string[] emailParts = PartnerContactEmail.Split('@');
            unique_id = unique_id + emailParts[1].Substring(0, 1);
        }

        // In one of our tests, an item was submitted with no answers selected in the SKU
        // information. This is not an expected path, so we need to write a system comment here.

        if (string.IsNullOrEmpty(SKUs))
        {
            SKUs = Microsoft.Operations.CSP.RegSys.Resources.No_SKU;
            //"[SYSTEM] No SKU items appear to have been selected in the form submission";
            // This information is potentially customer-facing! Keep this message generic.
            Warnings.Add("Expecting SKU Information, but no information was found. Observe the original XML fragment for diagnostic information.");
            Warnings.Add("This regsys entry will be processed, but is absent any SKU information. Recommend closing as 'cannot process'.");
        }

        VirtualKey = string.Format("{0:yyyyMMddHHmmss}{1}{2}", CreatedDateInRegSys, PartnerContactName.Substring(0, 1), unique_id).ToUpper();

        // PartnerGUID = VirtualKey; // commenting out this. Partner GUID is actually what's captured
        // via the RegSys form.
        OriginalFragment = xml.InnerXml;
    }

    public string AdditionalEmail { get; set; }

    public string AgreementToTerms { get; set; }

    public string BillingContactEmail { get; set; }

    public string BillingContactName { get; set; }

    public DateTime CreatedDateInRegSys { get; set; }

    public string CustomerTenantDomain { get; set; }

    /// <summary>
    /// Where did this information come from? e.g. URL (for Web Lead / Interest Form) or Filename
    /// (for spreadsheet)
    /// </summary>
    public string DataOrigin { get; set; }

    public string DataOriginDetail { get; set; }

    /// <summary>
    /// If the item is known to have an equivalent entry in TFS, you can
    /// </summary>
    public int ExistingTfsID { get; set; }

    public string IPAddress { get; set; }

    public string IsAuthenticated { get; set; }

    public string IsEmailVerified { get; set; }

    public string LCID { get; set; }

    public string MPNID { get; set; }

    public string OriginalFragment { get; set; }

    public string PartnerArea { get; set; }

    public string PartnerContactEmail { get; set; }

    public string PartnerContactName { get; set; }

    public string PartnerOrganizationName { get; set; }

    public string SKUs { get; set; }

    public string SpecialInstructions { get; set; }

    /// <summary>
    /// The RegSys data does not return a virtual key, so we have to derive one. We'll track the
    /// email response
    /// </summary>
    public string VirtualKey { get; set; }

    public string WizardGUID { get; set; }
}