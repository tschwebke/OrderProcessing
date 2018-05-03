using System.Globalization;

namespace Microsoft.Operations
{
    public static partial class Extensions
    {
        /// <summary>
        /// Simple test to see whether an object/input is purely numeric.
        /// </summary>
        public static bool IsNumeric(this object value)
        {
            double d;
            return double.TryParse(value.ToString(), System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture.NumberFormat, out d);
        }
    }
}