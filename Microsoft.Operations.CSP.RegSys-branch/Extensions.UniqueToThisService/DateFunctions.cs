using System;

namespace Microsoft.Operations.CSP.RegSys
{
    /// <summary>
    /// Variety of date-based functions which may be useful when working with Date values.
    /// </summary>
    public static class DateFunctions
    {
        /// <summary>
        /// Adds the given number of business days to the <see cref="DateTime"/>.
        /// </summary>
        /// <param name="current">The date to be changed.</param>
        /// <param name="days">Number of business days to be added.</param>
        /// <returns>A <see cref="DateTime"/> increased by a given number of business days.</returns>
        public static DateTime AddBusinessDays(this DateTime current, int days)
        {
            var sign = Math.Sign(days);
            var unsignedDays = Math.Abs(days);
            for (var i = 0; i < unsignedDays; i++)
            {
                do
                {
                    current = current.AddDays(sign);
                }
                while (current.DayOfWeek == DayOfWeek.Saturday ||
                    current.DayOfWeek == DayOfWeek.Sunday);
            }
            return current;
        }

        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = dt.DayOfWeek - startOfWeek;
            if (diff < 0)
            {
                diff += 7;
            }
            return dt.AddDays(-1 * diff).Date;
        }

        public static DateTime SubscriptionFinalDate(this DateTime current)
        {
            if (current.DayOfWeek != DayOfWeek.Monday && current.DayOfWeek != DayOfWeek.Tuesday)
            {
                current = current.AddDays(7);
                int daysUntilMonday = (DayOfWeek.Monday - current.DayOfWeek + 7) % 7;
                current = current.AddDays(daysUntilMonday);
            }
            else
            {
                if (current.DayOfWeek == DayOfWeek.Monday) { current = current.AddDays(1); }
                int daysUntilMonday = (DayOfWeek.Monday - current.DayOfWeek + 7) % 7;
                current = current.AddDays(daysUntilMonday);
            }
            return current;
        }

        /// <summary>
        /// Subtracts the given number of business days to the <see cref="DateTime"/>.
        /// </summary>
        /// <param name="current">The date to be changed.</param>
        /// <param name="days">Number of business days to be subtracted.</param>
        /// <returns>A <see cref="DateTime"/> increased by a given number of business days.</returns>
        public static DateTime SubtractBusinessDays(this DateTime current, int days)
        {
            return AddBusinessDays(current, -days);
        }
    }
}