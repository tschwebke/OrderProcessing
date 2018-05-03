using System;
using System.Data;
using System.Data.SqlClient;

static public partial class Extensions
{
    /// <summary>
    /// Specify the value that you want to use as input to the stored procedure (typically for an
    /// input parameter). This full definition will/should be the one which is used the most -
    /// BOOLEAN and INTEGER types
    /// </summary>
    /// <param name="parameterName">
    /// Name of the parameter as known to the stored procedure. This typically has a "@" in front of it.
    /// </param>
    /// <param name="value">
    /// Supply the value you wish to inject. Note: SQL Server may not accept null values.
    /// </param>
    /// <param name="dataType">
    /// The data type of the parameter, as specified in the stored procedure. For use with Input
    /// parameters, this is typically required.
    /// </param>
    /// <param name="value">Null can be passed as a value (accepted by Sql Server).</param>
    /// <remarks>
    /// For use with Input parameters, this is typically required. For use with Output parameters,
    /// the value is usually retrieved after the command is executed.
    /// </remarks>
    public static void AddInputParameter(this SqlCommand cmd, String parameterName, Object value, SqlDbType dataType)
    {
        AddParameter(cmd, ParameterDirection.Input, parameterName, dataType, value, 0, 0, 0);
    }

    /// <summary>
    /// Specifying the size for types of CHAR and VARCHAR is required.
    /// </summary>
    /// <param name="size">The length of the character.</param>
    public static void AddInputParameter(this SqlCommand cmd, String parameterName, Object value, SqlDbType dataType, int size)
    {
        AddParameter(cmd, ParameterDirection.Input, parameterName, dataType, value, size, 0, 0);
    }

    /// <summary>
    /// Specifying details for DECIMAL data type is required. In this case, the 'length' is
    /// automatically handled.
    /// </summary>
    /// <param name="precision">Total length of the decimal number, including decimal places</param>
    /// <param name="scale">Number of decimal places (i.e. after the ".")</param>
    public static void AddInputParameter(this SqlCommand cmd, String parameterName, Object value, SqlDbType dataType, byte precision, byte scale)
    {
        AddParameter(cmd, ParameterDirection.Input, parameterName, dataType, value, 0, precision, scale);
    }

    /// <summary>
    /// Default output parameter, for INTEGER and BOOLEAN types (where the size is already defined)
    /// </summary>
    public static void AddOutputParameter(this SqlCommand cmd, String parameterName, SqlDbType dataType)
    {
        AddParameter(cmd, ParameterDirection.Output, parameterName, dataType, null, 0, 0, 0);
    }

    /// <summary>
    /// Adds an OUTPUT PARAMETER
    /// </summary>
    public static void AddOutputParameter(this SqlCommand cmd, String parameterName, SqlDbType dataType, int size)
    {
        AddParameter(cmd, ParameterDirection.Output, parameterName, dataType, null, size, 0, 0);
    }

    /// <summary>
    /// For decimal data types, precision (and scale) are required (replaces the size definition)
    /// </summary>
    public static void AddOutputParameter(this SqlCommand cmd, String parameterName, SqlDbType dataType, byte precision, byte scale)
    {
        AddParameter(cmd, ParameterDirection.Output, parameterName, dataType, null, 0, precision, scale);
    }

    public static decimal GetDecimal(this SqlCommand command, String outputParameterName)
    {
        try
        {
            return Convert.ToDecimal(command.Parameters[outputParameterName].Value);
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// Attempts to return a integer-formatted variable from the SQL Server stored proc output
    /// variable. If it can't do it, then it returns a zero. This is effectively an easy catch-all
    /// for null values.
    /// </summary>
    public static int GetInteger(this SqlCommand command, String outputParameterName)
    {
        try
        {
            return Convert.ToInt32(command.Parameters[outputParameterName].Value);
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// Handy wrapper which just spits out a string version of whatever the parameter value was.
    /// </summary>
    public static string GetString(this SqlCommand command, String outputParameterName)
    {
        return Convert.ToString(command.Parameters[outputParameterName].Value);
    }

    /// <summary>
    /// Adds a parameter to the command object, using given parameters from the extension methods.
    /// Note the conditional logic, this is meant to handle most scenarios (but not necessarily all).
    /// Some elements are not required depending on the combination of all the other data.
    /// </summary>
    private static void AddParameter(SqlCommand command, ParameterDirection inOrOut, String parameterNameFromStoredProcedure, SqlDbType dataType, Object parameterValue, int size, byte decimalPrecision, byte decimalScale)
    {
        SqlParameter param = new SqlParameter();
        param.Direction = inOrOut;
        param.ParameterName = parameterNameFromStoredProcedure;
        param.Value = parameterValue;
        param.SqlDbType = dataType;

        if (size != 0)
        {
            param.Size = size;
        }

        if (dataType == SqlDbType.Decimal && decimalPrecision != 0 && decimalScale != 0)
        {
            param.Precision = decimalPrecision;
            param.Scale = decimalScale;
        }

        command.Parameters.Add(param);
    }
}