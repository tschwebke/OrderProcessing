using System;

public static partial class Extensions
{
    /// <summary>
    /// Useful for spreadsheets.
    /// Credit: Stackoverflow Internet Code
    /// </summary>
    public static string ToExcelColumnName(this int columnNumber)
    {
        int dividend = columnNumber;
        string columnName = string.Empty;
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }

        return columnName;
    }
}