using System.Collections;
using System.IO;

public static partial class Extensions
{
    /// <summary>
    /// Returns file names from given folder that comply to given filters
    /// </summary>
    /// <param name="SourceFolder">Folder with files to retrieve</param>
    /// <param name="Filter">Multiple file filters separated by | character</param>
    /// <param name="searchOption">File.IO.SearchOption, could be AllDirectories or TopDirectoryOnly</param>
    /// <returns>
    /// Array of FileInfo objects that presents collection of file names that meet given filter
    /// </returns>
    public static string[] GetFiles(string SourceFolder, string Filter, SearchOption searchOption)
    {
        // ArrayList will hold all file names
        ArrayList alFiles = new ArrayList();

        // Create an array of filter string
        string[] MultipleFilters = Filter.Split('|');

        // for each filter find mathing file names
        foreach (string FileFilter in MultipleFilters)
        {
            // add found file names to array list
            alFiles.AddRange(Directory.GetFiles(SourceFolder, FileFilter, searchOption));
        }

        // returns string array of relevant file names
        return (string[])alFiles.ToArray(typeof(string));
    }
}