using Microsoft.Operations.CSP.RegSys;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;

namespace Microsoft.Operations
{
    public static partial class TeamFoundationServerHelper
    {
        /// <summary>
        /// This routine will place the relevant SKU number IN FRONT of the text in the line item in
        /// the 'order'. IT ASSUMES THE TEXT HAS NOT BECOME MALFORMED through editing or otherwise,
        /// WHICH WOULD AFFECT THIS ROUTINE.
        /// </summary>
        public static void CleanSkuInformation(this WorkItem wi, bool includeHistoryWrite = false)
        {
            try
            {
                wi.PartialOpen();
                var formLang = wi.Fields["Microsoft.Operations.Partners.FormLanguage"].Value;

                XDocument documentLang = XDocument.Load(string.Format(@"{0}\\Languages.xml", FileSystem.ExecutingFolder));
                var displayLang =
                    (from c in documentLang.Root.Elements("LCID")
                     where (string)c.Attribute("language") == formLang.ToString() && c.Attribute("lang") != null
                     select c.Attribute("lang").Value).SingleOrDefault();

                Assembly assem = typeof(Service1).Assembly;
                ResourceHelper res = new ResourceHelper("Microsoft.Operations.CSP.RegSys.Resources", assem);

                string existingSkuInformation = wi.GetFieldValue("Microsoft.Operations.CreditDiscountApproval.Request.SKUs");
                string promoCodeExists = wi.GetFieldValue("Microsoft.Operations.Partners.Service.PromoCode");
                //existingSkuInformation = existingSkuInformation.Replace("(local taxes apply),5-Pack On Premises Support Incidents", "(local taxes apply)\n5-Pack On Premises Support Incidents");
                //existingSkuInformation = existingSkuInformation.Replace("31,800 円 (税別),5 インシデント パック", "31,800 円 (税別)\n5 インシデント パック");
                // Some malformed test entries do not have any SKU information contained. Do not
                // perform any work on these ones. NOTE: Later we'll have to mark these as being
                // invalid. Additionally, ignore any entries where the SKU information has already
                // been modified by the system

                if (!string.IsNullOrEmpty(existingSkuInformation)) //&& !existingSkuInformation.Contains("[SYSTEM]"
                {
                    string[] existing_order_lines = existingSkuInformation.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    List<string> cleaned_order_lines = new List<string>();

                    foreach (string item in existing_order_lines)
                    {
                        //var resKey = res.GetResourceNameContains(item, displayLang);
                        //Thread.CurrentThread.CurrentUICulture = new CultureInfo(displayLang);
                        //var resValue = res.GetResourceValue(resKey);

                        cleaned_order_lines.Add(item);
                    }

                    string cleanSkuInformation = string.Join(Environment.NewLine, cleaned_order_lines);
                    wi.Fields["Microsoft.Operations.CreditDiscountApproval.Request.SKUs"].Value = cleanSkuInformation;

                    if (wi.IsDirty && includeHistoryWrite)
                    {
                        wi.History += string.Format("<br/>[SYSTEM] Added actual SKU values (ref: Task 67509 for details and maintenance)");
                    }

                    // We don't actually do any saving here
                } // end of testing to see we actually have SKU information to work with
            }
            catch (Exception ex)
            {
                // Email Administrator, or Write to the Event log
            }
        }
    }
}