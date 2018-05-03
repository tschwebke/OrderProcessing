using System;
using System.Globalization;
using System.Reflection;

namespace Microsoft.Operations
{
    public static class ExtendAssembly
    {
        /// <summary>
        /// Returns information about the build/assembly version in #.#.#.# format.
        /// NOTE: This implementation is dependent on the 'Assembly' information being set to '1.0.*'
        /// If you notice that the build information is not displaying as expected, then check that first.
        /// </summary>
        public static string GetVersion(this Assembly x)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            if (asm.FullName != null)
            {
                string[] parts = asm.FullName.Split(',');
                string version = parts[1];

                long build = long.Parse(version.Split('.')[2]);
                double revision = double.Parse(version.Split('.')[3]);
                DateTime buildDate = new DateTime(2000, 1, 1).AddDays(build);
                buildDate = new DateTime(2000, 1, 1).AddDays(build).AddSeconds(revision * 2);
                return buildDate.ToString("yyyy.MM.dd.HHmm", CultureInfo.InvariantCulture);
            }
            else return string.Empty;
        }
    }
}