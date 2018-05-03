using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Microsoft.Operations
{
    public static partial class Extensions
    {
        /// <summary>
        /// usage: dInfo.GetFilesByExtensions(".jpg",".exe",".gif");
        /// </summary>
        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }
    }
}