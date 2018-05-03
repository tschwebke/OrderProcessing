using SevenZip;
using System.Collections.Generic;
using System.IO;
using System.Threading;

public static class SEVENZIP
{
    /// <summary>
    /// Decompress an archived file, returning the final output as a list, with the option to delete
    /// original file. IMPORTANT NOTE: Requires the 7s.dll file (included is the x86 version) to be
    /// accessible on target machine, otherwise fails with error. I tried, but I couldn't get this
    /// working 100% because I can't create a one-size-fits-all package. Will do manually for now. https://github.com/luuksommers/SevenZipSharp.Interop/
    /// </summary>
    public static List<string> UnSeven(string compressedFile, string targetDirectory = "", bool forceDeleteOriginalFile = false)
    {
        FileInfo fi = new FileInfo(compressedFile);
        if (string.IsNullOrEmpty(targetDirectory)) targetDirectory = fi.DirectoryName; // by default, use current folder location

        List<string> newFileNames = new List<string>();

        // To decompress:
        using (SevenZipExtractor sze = new SevenZipExtractor(fi.FullName))
        {
            sze.ExtractArchive(targetDirectory);
        }

        if (forceDeleteOriginalFile && newFileNames.Count > 0)
        {
            Thread.Sleep(500); // Wait to remove the lock, all objects to dispose
            fi.Delete();
        }

        return newFileNames;
    }
}