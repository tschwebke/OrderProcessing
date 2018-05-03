using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

/// <summary>
/// Common Class for handle ZIP operations.
/// </summary>
public static class ZIP
{
    /// <summary>
    /// Takes a file and compresses it.
    /// </summary>
    /// <param name="finalfile">name of the zip which</param>
    /// <param name="sourcefile">include directory</param>
    public static void CreateZipFile(string finalfile, string sourcefile)
    {
        try
        {
            using (ZipFile zip = new ZipFile())
            {
                zip.AddFile(sourcefile, string.Empty);
                zip.Save(finalfile);
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error occured: " + ex);
        }
    }

    /// <summary>
    /// To the same directory as the zip file
    /// </summary>
    public static List<string> Unzip(string existingZipFile, string targetDirectory = "", bool forceDeleteOriginalFile = false)
    {
        FileInfo fi = new FileInfo(existingZipFile);
        if (string.IsNullOrEmpty(targetDirectory)) targetDirectory = fi.DirectoryName; // by default, use current folder location

        List<string> newFileNames = new List<string>();

        using (ZipFile zip = ZipFile.Read(existingZipFile))
        {
            foreach (ZipEntry e in zip)
            {
                e.Extract(targetDirectory, ExtractExistingFileAction.OverwriteSilently);
                newFileNames.Add(e.FileName);
            }
        }

        if (forceDeleteOriginalFile && newFileNames.Count > 0)
        {
            Thread.Sleep(500); // Wait to remove the lock, all objects to dispose
            fi.Delete();
        }

        return newFileNames;
    }
}