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
                existingSkuInformation = existingSkuInformation.Replace("(local taxes apply),5-Pack On Premises Support Incidents", "(local taxes apply)\n5-Pack On Premises Support Incidents");
                existingSkuInformation = existingSkuInformation.Replace("31,800 円 (税別),5 インシデント パック", "31,800 円 (税別)\n5 インシデント パック");
                // Some malformed test entries do not have any SKU information contained. Do not
                // perform any work on these ones. NOTE: Later we'll have to mark these as being
                // invalid. Additionally, ignore any entries where the SKU information has already
                // been modified by the system

                if (!string.IsNullOrEmpty(existingSkuInformation)) //&& !existingSkuInformation.Contains("[SYSTEM]"
                {
                    // Apply some basic sanitation:
                    //1 - Pack On Premises Support Incident – USD 499(local taxes apply)
                    //1 インシデント – 49,900 円(税別)

                    //5 - Pack On Premises Support Incidents – USD 1,999(local taxes apply)
                    //5 インシデント パック – ¥159,900 円(税別)

                    //1 Cloud Consult Engagement – USD 2,000(local taxes apply)
                    //1 回 クラウド相談会 – 220,000 円(税別)

                    //20 Hours of Services Account Management – USD 4,000(local taxes apply)
                    //20 時間 サービス アカウント マネージャー追加時間 – 440,000 円(税別)

                    //Advanced Support for Partners Annual Subscription in Emerging Markets, One-Time Payment – USD 10,000 (local taxes apply).
                    //No Japanese translation

                    //Advanced Support for Partners Monthly Subscription – Monthly Payment – USD 1, 250(local taxes apply)
                    //Advanced Support for Partners 月額課金 – 毎月のお支払い – 137, 500 円(税別)

                    //Advanced Support for Partners Annual Subscription – One - Time Payment, 3 % discount – USD 14, 550(local taxes apply)
                    //Advanced Support for Partners 年間契約 – 一括のお支払いで 3 % のディスカウント – 1, 600, 500 円(税別)

                    //ASfP PA Transition Package - One - Time Payment

                    //Offer catalog               USD JPY     SKU
                    //ASfP Monthly Subscription		$1, 250	¥137, 500    W6M - 00001
                    //ASfP Annual Subscription		$15, 000	¥1, 650, 000
                    //ASfP Annual Subscription(3 % discount)	$14, 550	¥1, 600, 500  W6N - 00001
                    //20 Hours SAM 				$4, 000	¥440, 000    W73 - 00001
                    //1 Cloud Consult Engagement		$2, 000	¥220, 000    W74 - 00001
                    //1 - Pack On Premises Support Incident	$499	¥31, 000     W67 - 00001
                    //5 - Pack On Premises Support Incidents	$1, 999	¥159, 000    W69 - 00001
                    //ASfP PA Transition Package - One - Time Payment           AAA - 13751

                    string[] existing_order_lines = existingSkuInformation.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    List<string> cleaned_order_lines = new List<string>();

                    foreach (string item in existing_order_lines)
                    {
                        var resKey = res.GetResourceNameContains(item, displayLang);
                        Thread.CurrentThread.CurrentUICulture = new CultureInfo(displayLang);
                        var resValue = res.GetResourceValue(resKey);

                        if (promoCodeExists == "")
                        {
                            if (resKey.Contains("SKU_Error") || resKey.Contains("No_SKU"))
                                cleaned_order_lines.Add("[SYSTEM] No SKU items appear to have been selected in the form submission");
                            else
                                cleaned_order_lines.Add(resValue);
                        }
                        else
                        {
                            if (resKey.Contains("SKU_Error") || resKey.Contains("ASfP_Monthly") || resKey.Contains("ASfP_Annual") || resKey.Contains("No_SKU") || resKey.Contains("ASfP_Emerging"))
                                cleaned_order_lines.Add(res.GetResourceValue("ASfP_PA_Transition"));
                            else
                                cleaned_order_lines.Add(resValue);
                        }
                    }

                    if (promoCodeExists != "")
                    {
                        Boolean pa = false;
                        foreach (string item in cleaned_order_lines)
                        {
                            if (item.Contains("ASfP PA Transition Package")) { pa = true; }
                        }
                        if (!pa) { cleaned_order_lines.Add("AAA-13751 ASfP PA Transition Package - One-Time Payment "); }
                    }

                    //if (existing_order_lines.Length != cleaned_order_lines.Count)
                    //{
                    //    cleaned_order_lines.Add("Unrecognized SKUs");
                    //    foreach (string item in existing_order_lines)
                    //    {
                    //        cleaned_order_lines.Add("Unrecognized SKU " + item);
                    //    }
                    //    wi.History += "<br/>[SYSTEM] Attempted to provide SKU codes, but some entries were not found.";
                    //}

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