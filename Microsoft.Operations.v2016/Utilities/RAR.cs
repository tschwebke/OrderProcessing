using SharpCompress.Common;
using SharpCompress.Reader;
using System.Collections.Generic;
using System.IO;
using System.Threading;

/// <summary>
/// Handle RAR compressed files, using 3rd party library. (our wrapper)
/// </summary>
public static class RAR
{
    /// <summary>
    /// Decompress an archived file, returning the final output as a list, with the option to delete
    /// original file.
    /// NOTE: Because this library handles multiple formats, we may make the naming more generic later.
    /// </summary>
    public static List<string> Unrar(string compressedFile, string targetDirectory = "", bool forceDeleteOriginalFile = false)
    {
        FileInfo fi = new FileInfo(compressedFile);
        if (string.IsNullOrEmpty(targetDirectory)) targetDirectory = fi.DirectoryName; // by default, use current folder location

        List<string> newFileNames = new List<string>();

        using (Stream stream = File.OpenRead(fi.FullName))
        {
            var reader = ReaderFactory.Open(stream);
            while (reader.MoveToNextEntry())
            {
                if (!reader.Entry.IsDirectory)
                {
                    newFileNames.Add(reader.Entry.FilePath);
                    reader.WriteEntryToDirectory(targetDirectory, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                }
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