using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using WebMarkupMin.Core.Minifiers;

/// <summary>
/// Optimization library for BIOS Operations.
/// </summary>
public static class Optimize
{
    /// <summary>
    /// Wrapper around the library known as 'WebMarkupMin' (a codeplex project, see Documentation
    /// folder for license). Takes a HTML stream and reduces it, but everything still works. Use this
    /// if you want to optimize and have no need for human-readability (e.g. html emails). Allows you
    /// to KEEP all HTML comments inside the HTML you're working with (for readability and sanity),
    /// safe in the knowledge that it won't be exposed later. Other internal options do exist if you
    /// need more flexibility. Please note: your mileage may vary: in some rare cases this library
    /// will actually make the HTML larger &gt; make sure you run some comparison tests.
    /// WARNING: May sometimes add a 'mailto:'
    /// </summary>
    /// <param name="htmlInput">your candidate input stream</param>
    /// <param name="compressionSummary">
    /// some statistics about the compression which occured (mostly for developer information)
    /// </param>
    /// <param name="includeContentInStatistics">
    /// if you need to see the final result, this flag will force it to be included in the statistics
    /// information (may save you some time)
    /// </param>
    /// <returns>HTML string without whitespace</returns>
    public static string MinifyHtml(string htmlInput, out string compressionSummary, bool includeContentInStatistics = false)
    {
        HtmlMinifier htmlMinifier = new HtmlMinifier();
        StringBuilder stats = new StringBuilder();

        MarkupMinificationResult result = htmlMinifier.Minify(htmlInput, generateStatistics: true);
        if (result.Errors.Count == 0)
        {
            MinificationStatistics statistics = result.Statistics;
            if (statistics != null)
            {
                stats.AppendLine(string.Format("HTML Minification - Took: {0} Milliseconds", statistics.MinificationDuration));
                stats.AppendLine(string.Format("Original size: {0:N0} Bytes", statistics.OriginalSize));
                stats.AppendLine(string.Format("Minified size: {0:N0} Bytes", statistics.MinifiedSize));
                stats.AppendLine(string.Format("Saved: {0:N2}%", statistics.SavedInPercent));
            }

            if (includeContentInStatistics)
            {
                stats.AppendLine(string.Format("Minified content:{0}{0}{1}", Environment.NewLine, result.MinifiedContent));
            }
        }
        else
        {
            IList<MinificationErrorInfo> errors = result.Errors;

            stats.AppendLine(string.Format("Found {0:N0} error(s):", errors.Count));
            stats.AppendLine(Environment.NewLine);

            foreach (var error in errors)
            {
                stats.AppendLine(string.Format("Line {0}, Column {1}: {2}", error.LineNumber, error.ColumnNumber, error.Message));
                stats.AppendLine();
            }
        }

        compressionSummary = stats.ToString();
        return result.MinifiedContent;
    }

    /// <summary>
    /// Use this for easy handling on all sorts of data which we expect to be string but may be all
    /// manner of null objects.
    /// </summary>
    /// <param name="input">The object you suspect has some data in it</param>
    /// <param name="stripHtml">Optionally force out any HTML tags, if there was a string value</param>
    public static string SafeString(object input, bool stripHtml = false)
    {
        string output = string.Empty;

        if (input != null && !DBNull.Value.Equals(input)) // go ahead to process
        {
            output = input.ToString(); // raw, to start with

            if (stripHtml)
            {
                output = Regex.Replace(output, "<.*?>", string.Empty); // remove all <div> and other formatting tags that SharePoint seems to always put in.
                output = output.Replace("\r\n", string.Empty).Trim(); // and any funny carriage returns and whitespace (SharePoint again)
                output = output.Replace("&#160;", string.Empty); // and other special characters (this is a non-breaking space)
            }
        }

        return output;
    }

    // TODO: http://www.codeproject.com/Tips/323212/Accurate-way-to-tell-if-an-assembly-is-compiled-in
}