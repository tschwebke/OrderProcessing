using System.Collections;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Threading;

namespace Microsoft.Operations.CSP.RegSys
{
    public class ResourceHelper
    {
        /// <summary>
        /// ResourceHelper
        /// </summary>
        /// <param name="resourceName">i.e. "Namespace.ResourceFileName"</param>
        /// <param name="assembly">i.e. GetType().Assembly if working on the local assembly</param>
        public ResourceHelper(string resourceName, Assembly assembly)
        {
            ResourceManager = new ResourceManager(resourceName, assembly);
        }

        private ResourceManager ResourceManager { get; set; }

        public string GetResourceName(string value, string culture)
        {
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(culture);
            DictionaryEntry entry = ResourceManager.GetResourceSet(Thread.CurrentThread.CurrentUICulture, true, true).OfType<DictionaryEntry>().FirstOrDefault(dictionaryEntry => dictionaryEntry.Value.ToString().Equals(value));
            return entry.Key.ToString();
        }

        public string GetResourceNameContains(string value, string culture)
        {
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(culture);
            DictionaryEntry entry = ResourceManager.GetResourceSet(Thread.CurrentThread.CurrentUICulture, true, true).OfType<DictionaryEntry>().FirstOrDefault(dictionaryEntry => dictionaryEntry.Value.ToString().Contains(value));
            return entry.Key.ToString();
        }

        public string GetResourceValue(string name)
        {
            string value = ResourceManager.GetString(name);
            return !string.IsNullOrEmpty(value) ? value : null;
        }
    }
}