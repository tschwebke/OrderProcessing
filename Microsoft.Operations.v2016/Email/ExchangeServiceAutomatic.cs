using Microsoft.Exchange.WebServices.Data;
using System;
using System.Runtime.Caching;

namespace Microsoft.Operations
{
    /// <summary>
    /// Extension for EWS service which does a bunch of stuff so you don't need complex calling code.
    /// Note: this does hide a fair amount of functionality, don't use it if you need full exposure
    ///       on the EWS components.
    ///
    /// NOTES:
    ///
    /// 1. This mechanic caches the Autodiscover Uri value, because although incredibly good and
    ///    wholesome, the Autodiscover method is expensive and slow so use this it will be accurate
    ///    and help performance a little.
    /// Autodiscover: http://blogs.technet.com/b/exchange/archive/2008/08/08/3406026.aspx Technical
    ///               Reference:
    ///               http://msdn.microsoft.com/EN-US/library/office/dn659837%28v=exchg.150%29.aspx
    ///               Also note: The cached value will only last as long as the MEMORY PROCESS for
    ///               this item. If used for websites and you experience performance problems with
    ///               sending, consider extending the expiry time on application pools, i.e. the
    ///               default is only active for 20 minutes which is great if you have a lot of
    ///               emails on a 'hot' service, but not so beneficial for sporadic usage. The actual
    ///               cache time is set at the recommended 24 hours, but you can force this on a hot
    ///               service by recycling the application pool.
    ///
    /// 2. The value specified for ExchangeVersion is hardcoded here (e.g. 'Exchange2010_SP1'),
    ///    because if not specified then the default is used which is most often incorrect. The EWS
    ///    usage is a bit misleading &gt;&gt; the value is there because it dictates the formatting
    ///    for the SOAP messages which get SENT. For most of the time, if you choose a recent
    ///    Exchange Version, for the most part it should be compatible with the target Exchange Server.
    /// BUT: This compatibility gets eroded over time as Exchange Versions come and go, so every few
    ///      years this may need some fine-tuning.
    ///
    /// 3. This class was built in response to the change of system accounts moving to the cloud
    ///    environment (ref: Task 4102). It's designed to reduce the inline code required for using
    ///    the EWS object, i.e. initialize using one line and the email address. Originally we used
    ///    "https://aps.mail.microsoft.com/ews/exchange.asmx" for the URL value, which was hardcoded
    ///    in most of our assemblies ;P But with the new cloud migration it should have this value
    ///    most of the time: "https://apj.cloudmail.microsoft.com/ews/exchange.asmx".
    /// </summary>
    public static class ExchangeServiceAutomatic
    {
        private static ObjectCache cache = MemoryCache.Default;

        // Please note, this cache is shared at the machine level. if you have a requirement for a
        // protected (app only) then initialize using MemoryCache("myUniqueRef").

        /// <summary>
        /// Creates an ExchangeService object which already has all the settings you're likely to encounter.
        /// 1. Uri is set for you automatically (using autodiscover + cache combination)
        /// 2. RequestedServerVersion is set the default for this assembly (may need adjusting on
        ///    future Exchange Versions).
        /// 3. The security context uses the current windows identity by default - see notes.
        /// </summary>
        /// <param name="emailAddress">
        /// Full SMTP address of the account nominated for mailbox &gt;&gt; which helps find the
        /// mailbox location. Please note the final connection will actually use the security context
        /// of the connecting thread! If they are not the same, you might get strange results with
        /// the email sending.
        /// </param>
        public static ExchangeService New(string emailAddress)
        {
            return New(emailAddress, null);
        }

        /// <summary>
        /// This overload allows you to supply an alternate WebCredentials identity. Use with TESTING
        /// ONLY, since it is not best practice to have passwords embedded in the code!
        /// </summary>
        /// <param name="emailAddress">Full SMTP address of the account you will be using.</param>
        /// <param name="accountPassword">
        /// Supplying this value will force the service to be in the security context of that
        /// account, not supplying it will force to use the executing thread.
        /// </param>
        public static ExchangeService New(string emailAddress, string accountPassword = "")
        {
            // First, check the cached key value of MAILBOX location, which will be used for this account.

            string keyName = "ews_url_" + emailAddress;
            Uri existing_ews_url = (Uri)cache[keyName];

            ExchangeService ews = new ExchangeService(ExchangeVersion.Exchange2013);

            if (!string.IsNullOrEmpty(accountPassword))
            {
                ews.Credentials = new WebCredentials(emailAddress, accountPassword);
                // ews.Credentials = new NetworkCredential("svcspsub", accountPassword, "REDMOND");
            }
            else
            {
                // by default, this would attempt to connect using the current security context. Be
                // warned however ... in the current cloud environment this is not working well. i.e.
                // sometimes the service account may not have the right cookies setup required for
                // the federated identity management to work correctly. we're still trying to iron
                // this out. sorry.
            }

            // Now set the Url location for the mailbox, if known from previous (cached for 24 hours,
            // assuming memory is kept alive) if it's not, then we need to use the Autodiscover methods,

            if (existing_ews_url != null)
            {
                ews.Url = existing_ews_url;
            }
            else
            {
                ews.AutodiscoverUrl(emailAddress, RedirectionCallback);
                // ews.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                if (ews.Url != null)
                {
                    cache.Add(keyName, ews.Url, DateTime.Now.AddMinutes(30));
                }
            }

            return ews;
        }

        /// <summary>
        /// Usage is described in the Exchange Web Services document, but it is way confusing.
        /// Essentially, it's a way that any redirections can be verified. This will occur most often
        /// with any REDMOND domain addresses, where more than one server may be in operation.
        /// </summary>
        private static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }
    }
}