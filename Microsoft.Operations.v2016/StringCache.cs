using System;
using System.Web.Caching;

namespace Microsoft.Operations
{
    /// <summary>
    /// Object which can help store a simple string value in the HttpRuntime cache (note: has a
    /// dependency on System.Web assembly)
    /// </summary>
    public class StringCache
    {
        private static Cache objectCache = System.Web.HttpRuntime.Cache;

        /// <summary>
        /// Obtains a string value from the cache &gt; if it exists and is valid. Not being there
        /// doesn't mean it wasn't set ... it may have already expired.
        /// </summary>
        public static string Read(string cacheKey)
        {
            string cacheValue = string.Empty;

            if (objectCache.Get(cacheKey) != null)
            {
                cacheValue = (String)objectCache.Get(cacheKey);
            }
            return cacheValue;
        }

        /// <summary>
        /// Places a nominated key/value pair into the cache, with no set expiry.
        /// </summary>
        public static void Write(string cacheKey, String cacheValue)
        {
            Write(cacheKey, cacheValue, 0);
        }

        /// <summary>
        /// Caches an object, with a specific duration in Seconds. If you don't want the object to
        /// expire, set an expiry of Zero
        /// </summary>
        public static void Write(string cacheKey, String cacheValue, double duration)
        {
            if (duration == 0)
            {
                objectCache.Insert(cacheKey, cacheValue);
            }
            else
            {
                objectCache.Insert(cacheKey, cacheValue, null, System.DateTime.Now.AddSeconds(duration), System.Web.Caching.Cache.NoSlidingExpiration);
            }
        }
    }
}