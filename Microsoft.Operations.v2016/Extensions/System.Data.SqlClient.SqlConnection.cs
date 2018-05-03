using System.Data;
using System.Data.SqlClient;

static public partial class ExtensionMethods
{
    /// <summary>
    /// Shortcut for bulk upload into table which has the same column names/structure as the
    /// datatable supplied. Only use this if you know the schemas are identical (typically
    /// dynamically determined).
    /// </summary>
    public static int BulkUploadData(this SqlConnection connection, DataTable dataTableObjectWithData, string destinationTableName)
    {
        using (SqlBulkCopy bulk_copy = new SqlBulkCopy(connection, SqlBulkCopyOptions.KeepIdentity, null))
        {
            // do the column mappings to bypass the identity column.
            foreach (DataColumn col in dataTableObjectWithData.Columns)
            {
                bulk_copy.ColumnMappings.Add(new SqlBulkCopyColumnMapping(col.ColumnName, col.ColumnName));
            }

            bulk_copy.BulkCopyTimeout = 0; // no timeout value
            bulk_copy.BatchSize = 1000;
            bulk_copy.DestinationTableName = destinationTableName;
            bulk_copy.WriteToServer(dataTableObjectWithData);
        }

        return dataTableObjectWithData.Rows.Count;
    }

    /// <summary>
    /// Generic one-liner for executing command against database
    /// </summary>
    public static void ExecuteNonQuery(this SqlConnection connection, string tsql)
    {
        using (var cmd = new SqlCommand())
        {
            cmd.Connection = connection;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = tsql;
            if (connection.State != ConnectionState.Open) connection.Open();
            cmd.ExecuteNonQuery();
        }
    }

    /// <summary>
    /// Forces an open connection, using automatically constructed connection string (uses current
    /// security context)
    /// </summary>
    public static void OpenConnectionToSqlServer(this SqlConnection connection, string serverName, string databaseName)
    {
        if (connection.State != ConnectionState.Open)
        {
            connection.ConnectionString = string.Format(@"Server={0};Database={1};Trusted_Connection=True;Connection Timeout=500", serverName, databaseName);
            connection.Open();
        }
    }
}