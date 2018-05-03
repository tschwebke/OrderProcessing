using System;
using System.Globalization;

namespace Microsoft.Operations
{
    /// <summary>
    /// Use to insert safe types into the spreadsheet when using EPPLUS (OpenOfficeXml)
    /// </summary>
    public static class SpreadsheetMagic
    {
        /// <summary>
        /// TODO: This needs some serious work ...
        /// </summary>
        /// <param name="possibleDateInput"></param>
        /// <returns></returns>
        public static object DateValue(object possibleDateInput)
        {
            DateTime possible;

            // DateTime.ParseExact(requestDateString.ToString(), @"M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
            //try
            //{
            //    DateTime.TryParseExact
            //    bool successful = true;
            //}

            // bool successful = DateTime.TryParseExact(possibleDateInput.ToString(), @"M/d/yyyy
            // h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out
            // possible); // american ...
            bool successful = DateTime.TryParseExact(possibleDateInput.ToString(), @"d/M/yyyy h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out possible); // on test server, australian.

            if (successful)
                return possible;
            else
                return string.Empty;
        }

        /// <summary>
        /// Uses a variety of date formatting options to get the date. WHY? Because In some
        /// situations, user may change the formatting of the spreadsheet :( The best is always where
        /// an OADate is used, however this isn't always available so the date value may need to be
        /// read in some other fashion. This attempts to cover most situations.
        /// </summary>
        /// <param name="otherExpectedFormat">
        /// e.g. 'MM/dd/yy' or 'MM/dd/yyyy' or whatever you're expecting the string value to be
        /// formatted as.
        /// </param>
        public static DateTime GetPossibleDate(object cellValue, string otherExpectedFormat)
        {
            DateTime output;

            try
            {
                output = DateTime.FromOADate(Convert.ToDouble(cellValue)); // NORMAL DATE FORMATTED
            }
            catch
            {
                DateTime.TryParseExact(cellValue.ToString(), otherExpectedFormat, System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out output);
            }
            finally
            {
            }

            return output;
        }
    }
}