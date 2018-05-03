using System;
using System.Globalization;
using System.Linq;

namespace Microsoft.Operations
{
    public static partial class ExtensionMethods
    {
        /// <summary>
        /// Use for adding business days only, by excluding specific days from the count.
        /// </summary>
        /// <param name="dayCount">The expected number of days you want to add</param>
        /// <param name="notIncludedDays">Specify which days do NOT count</param>
        public static DateTime AddDays(this DateTime from, int dayCount, params DayOfWeek[] notIncludedDays)
        {
            DateTime projectedDate = from;
            int totalCalendarDays = 0;
            int dayCountvirtual = 0;

            for (int i = 1; i <= dayCount * 2; i++)
            {
                if (!notIncludedDays.Contains(projectedDate.AddDays(i).DayOfWeek))
                {
                    dayCountvirtual++;
                }

                totalCalendarDays++; // but always increment

                if (dayCount == dayCountvirtual) break;
            }

            return projectedDate.AddDays(totalCalendarDays);
        }

        public static double AgeInMinutes(this DateTime from)
        {
            TimeSpan ts = DateTime.Now - from;
            return ts.TotalMinutes;
        }

        // Very quick syntax for seeing if a specific date lies between to other dates Exact matches
        // on either start or finish = true;
        public static bool Between(this DateTime midpoint, DateTime start, DateTime finish)
        {
            if (midpoint <= finish && midpoint >= start)
                return true;
            else
                return false;
        }

        public static int GetWeekOfMonth(this DateTime time)
        {
            DateTime first = new DateTime(time.Year, time.Month, 1);
            return time.GetWeekOfYear() - first.GetWeekOfYear() + 1;
        }

        /// <summary>
        /// Determines whether or not the sample time falls into Mon-Fri 9.00am to 5.00pm
        /// </summary>
        /// <param name="startHour">change if you need something different than 9.00</param>
        public static bool IsBusinessHours(this DateTime date, int startHour = 9, int finishHour = 17)
        {
            return (date.TimeOfDay.Hours >= startHour && date.TimeOfDay.Hours <= finishHour && date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday) ? true : false;
        }

        /// <summary>
        /// Similar to the Visual Basic method, gives the month name in full. (but without needing a
        /// reference to the Visual Basic DLL). Not very sophisticated, but it does get used a lot in reports.
        /// </summary>
        public static string MonthName(this DateTime value)
        {
            int month = value.Month;

            switch (month)
            {
                case 1: return "January";
                case 2: return "February";
                case 3: return "March";
                case 4: return "April";
                case 5: return "May";
                case 6: return "June";
                case 7: return "July";
                case 8: return "August";
                case 9: return "September";
                case 10: return "October";
                case 11: return "November";
                case 12: return "December";
                default: return string.Empty;
            }
        }

        public static DateTime Next(this DateTime from, DayOfWeek dayOfWeek)
        {
            int start = (int)from.DayOfWeek;
            int target = (int)dayOfWeek;
            if (target <= start)
                target += 7;
            return from.AddDays(target - start);
        }

        /// <summary>
        /// Convert a UTC Date to a named TimeZone to find out the local time in that timezone
        /// TODO: Merge this with the equivalent in 'Date Magic'
        /// </summary>
        /// <param name="utcDate">A UTC date that you want to use</param>
        /// <param name="timeZoneName">e.g. "Eastern Standard Time" or "Pacific Standard Time"</param>
        public static DateTime ToTimeInZone(this DateTime utcDate, string timeZoneName)
        {
            TimeZoneInfo tz = TimeZoneInfo.FindSystemTimeZoneById(timeZoneName);
            return TimeZoneInfo.ConvertTimeFromUtc(utcDate, tz);
        }

        private static int GetWeekOfYear(this DateTime time)
        {
            return new GregorianCalendar().GetWeekOfYear(time, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
        }
    }
}