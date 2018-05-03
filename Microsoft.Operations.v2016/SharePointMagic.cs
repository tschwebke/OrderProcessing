using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace Microsoft.Operations
{
    /// <summary>
    /// Sometimes we have to support Microsoft SharePoint. This collection of methods is designed to
    /// help make that process more bearable.
    /// </summary>
    public static class SharePointMagic
    {
        /// <summary>
        /// Takes a SharePoint object and turns it into a string. Also strips away any HTML that
        /// might be associated there, since the Sharepoint formatting contains all manner of Junk.
        /// </summary>
        public static string SafeString(object input)
        {
            string output = string.Empty;

            if (input != null) // go ahead to process
            {
                output = input.ToString(); // to start
                output = Regex.Replace(output, "<.*?>", string.Empty); // remove all <div> and other formatting tags that SharePoint seems to always put in.
                output = output.Replace("\r\n", string.Empty).Trim(); // and any funny carriage returns and whitespace (SharePoint again)
                output = output.Replace("&#160;", string.Empty); // and other special characters (this is a non-breaking space)
            }

            return output;
        }

        /// <summary>
        /// Determines (properly, using object model) whether
        /// </summary>
        /// <param name="context">
        /// The client context should already be a connection to the expected root parent.
        /// </param>
        /// <param name="siteUrl"></param>
        /// <param name="expectedChildWebUrl"></param>
        /// <returns></returns>
        public static bool SiteExists(ClientContext context, string expectedChildWebUrl)
        {
            // load up the root web object but only specifying the sub webs property to avoid
            // unneeded network traffic
            var web = context.Web;
            context.Load(web, w => w.Webs);
            context.ExecuteQuery();
            // use a simple linq query to get any sub webs with the URL we want to check
            var subWeb = (from w in web.Webs where w.Url == expectedChildWebUrl select w).SingleOrDefault();
            if (subWeb != null)
            {
                // if found true
                return true;
            }
            // default to false...
            return false;
        }

        /// <summary>
        /// This variant uses http request. Don't use because very inefficient - it downloads the
        /// entire repsonse across network.
        /// </summary>
        private static bool SiteExistsSlowVersion(string url, System.Net.ICredentials credentials)
        {
            try
            {
                Uri uri = new Uri(url);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                request.Credentials = credentials;

                request.Method = "GET";
                string result = "";
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (Stream responseStream = response.GetResponseStream())
                    {
                        using (StreamReader readStream = new StreamReader(responseStream, System.Text.Encoding.UTF8))
                        {
                            result = readStream.ReadToEnd();
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ListWebTemplates()
        // Create, using the template :(
        //WebTemplateCollection webTemplates = context.Web.GetAvailableWebTemplates(1033, true);
        //context.Load(webTemplates);
        //context.ExecuteQuery();
    }
}