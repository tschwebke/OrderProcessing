using System;
using System.Collections.Generic;

/// <summary>
/// Variety of date-based functions which may be useful when working with Date values.
/// </summary>
public static class DateMagic
{
    //// TODO: This needs to converted to extension methods (just a bit easier!)

    /// <summary>
    /// Calculates number of business days, taking into account:
    /// - weekends (Saturdays and Sundays)
    /// - bank holidays in the middle of the week
    /// </summary>
    /// <param name="firstDay">First day in the time interval</param>
    /// <param name="lastDay">Last day in the time interval</param>
    /// <param name="knownHolidays">List of known holidays which aren't business days (optional)</param>
    /// <returns>Number of business days during the 'span'</returns>
    public static int BusinessDaysBetween(this DateTime firstDay, DateTime lastDay, List<DateTime> knownHolidays = null)
    {
        firstDay = firstDay.Date;
        lastDay = lastDay.Date;
        if (firstDay > lastDay)
            throw new ArgumentException("Incorrect last day " + lastDay);

        TimeSpan span = lastDay - firstDay;
        int businessDays = span.Days + 1;
        int fullWeekCount = businessDays / 7;
        // find out if there are weekends during the time exceedng the full weeks
        if (businessDays > fullWeekCount * 7)
        {
            // we are here to find out if there is a 1-day or 2-days weekend in the time interval
            // remaining after subtracting the complete weeks
            int firstDayOfWeek = firstDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)firstDay.DayOfWeek;
            int lastDayOfWeek = lastDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)lastDay.DayOfWeek;

            if (lastDayOfWeek < firstDayOfWeek)
                lastDayOfWeek += 7;
            if (firstDayOfWeek <= 6)
            {
                if (lastDayOfWeek >= 7)// Both Saturday and Sunday are in the remaining time interval
                    businessDays -= 2;
                else if (lastDayOfWeek >= 6)// Only Saturday is in the remaining time interval
                    businessDays -= 1;
            }
            else if (firstDayOfWeek <= 7 && lastDayOfWeek >= 7)// Only Sunday is in the remaining time interval
                businessDays -= 1;
        }

        // subtract the weekends during the full weeks in the interval
        businessDays -= fullWeekCount + fullWeekCount;

        //// subtract the number of bank holidays during the time interval
        if (knownHolidays != null)
            foreach (DateTime bankHoliday in knownHolidays)
            {
                DateTime bh = bankHoliday.Date;
                if (firstDay <= bh && bh <= lastDay)
                    --businessDays;
            }

        return businessDays;
    }

    /// <summary>
    /// SHORTHAND For use with translating DateTime? into values which can be used with database directly
    /// </summary>
    public static object DateOrDbNull(DateTime? suppliedDate)
    {
        if (suppliedDate != null)
        {
            return (DateTime)suppliedDate;
        }
        else
        {
            return DBNull.Value;
        }
    }

    public static string DifferenceAsDays(DateTime? possibleDateFrom, bool returnEmptyIfNoValue = true)
    {
        if (possibleDateFrom == null)
        {
            return (returnEmptyIfNoValue) ? string.Empty : "no date";
        }
        else
        {
            return DifferenceAsDays(Convert.ToDateTime(possibleDateFrom));
        }
    }

    /// <summary>
    /// TODO: This is sloppy and needs work ! (i.e. work on the parameters and other options like pluralization).
    /// </summary>
    /// <remarks>Original usage/design was just in Microsoft.Operations.Alerts.Compliance</remarks>
    public static string DifferenceAsDays(DateTime possibleDateFrom)
    {
        // Find the number of days difference, express it as an absolute number

        int dayCount = Convert.ToInt32((Convert.ToDateTime(possibleDateFrom) - DateTime.Now).TotalDays);
        string plural = (System.Math.Abs(dayCount) > 1) ? "s" : string.Empty;

        if (dayCount == 0)
        {
            return string.Format("today"); // based on context
        }
        else if (dayCount < 0)
        {
            return string.Format("{0} day{1} ago", System.Math.Abs(dayCount), plural);
        }
        else
        {
            return string.Format("{0} day{1}", System.Math.Abs(dayCount), plural);
        }
    }

    public static string DifferenceAsDays(string preceding_text, DateTime possibleDateFrom)
    {
        return string.Format("{0} {1}", preceding_text, DifferenceAsDays(possibleDateFrom));
    }

    /// <summary>
    /// Shorthand for the first day of the month
    /// </summary>
    public static DateTime FirstDayOfMonth(DateTime dateTime)
    {
        return new DateTime(dateTime.Year, dateTime.Month, 1);
    }

    /// <summary>
    /// This version supports Sat/Sun only.
    /// </summary>
    public static bool IsLastBusinessDayInMonth(this DateTime sampleDate)
    {
        return (sampleDate.Date == LastBusinessDayInMonth(sampleDate.Year, sampleDate.Month)) ? true : false;
    }

    // credit: http://stackoverflow.com/questions/273048/how-to-determine-the-last-business-day-in-a-given-month
    public static DateTime LastBusinessDayInMonth(int year, int month)
    {
        var lastDay = DateTime.DaysInMonth(year, month);
        return PreviousOrCurrentBusinessDay(new DateTime(year, month, lastDay));
    }

    /// <summary>
    /// Determines the last day of the month Will work even with leap year dates.
    /// </summary>
    public static DateTime LastDayOfMonth(DateTime dateTime)
    {
        DateTime firstDayOfTheMonth = new DateTime(dateTime.Year, dateTime.Month, 1);
        return firstDayOfTheMonth.AddMonths(1).AddDays(-1);
    }

    public static DateTime PreviousOrCurrentBusinessDay(DateTime? beforeOrOnDate = null)
    {
        var fromDate = beforeOrOnDate ?? DateTime.Today;
        var year = fromDate.Year;
        var month = fromDate.Month;
        var day = fromDate.Day;
        var dtCurrent = new DateTime(year, month, day);

        while (!(dtCurrent.DayOfWeek < DayOfWeek.Saturday && dtCurrent.DayOfWeek > DayOfWeek.Sunday))
        {
            dtCurrent = dtCurrent.AddDays(-1);
        }
        return dtCurrent;
    }

    /// <summary>
    /// Basic routine for reading a probably date value from some source. Fails silently will give
    /// null returned. If you need sth more advanced, try other date conversion routines, this is
    /// just a basic shortcut only
    /// </summary>
    public static DateTime? ToDateTimeNullable(object input)
    {
        DateTime? output = null;

        try
        {
            output = Convert.ToDateTime(input);
        }
        catch
        {
            // some conversion error
        }

        return output;
    }

    /// <summary>
    /// Shorthand for getting date in local time, INCLUDING an allowance for daylight savings.
    /// </summary>
    /// <param name="zoneByName">
    /// Named zone, e.g. "Eastern Standard Time" (new york), "Pacific Standard Time" (redmond),
    /// "Tokyo Standard Time" or whatever applies
    /// </param>
    public static DateTime ToLocalTime(this DateTime utcDate, string zoneByName)
    {
        TimeZoneInfo tz = TimeZoneInfo.FindSystemTimeZoneById(zoneByName);
        return TimeZoneInfo.ConvertTimeFromUtc(utcDate, tz);
    }
}