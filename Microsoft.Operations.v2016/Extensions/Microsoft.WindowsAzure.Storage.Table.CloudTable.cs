using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;

public static class CloudTableExtensions
{
    /// <summary>
    /// Shorthand for using date in ####.##.##.##.## format.
    /// </summary>
    //public static T Read<T>(this CloudTable table, string partitionKey, DateTime date)
    //{
    //    return table.Read<T>(partitionKey, date.AsRowKey());
    //}

    ///// <summary>
    ///// Direct retrieval of key, object based on target table.
    ///// Returns null if not present ... so be sure to catch.
    ///// </summary>
    //public static T Read<T>(this CloudTable table, string partitionKey, string rowKey)
    //{
    //    TableOperation retrieveOperation = TableOperation.Retrieve<T>(partitionKey, rowKey);
    //    TableResult retrievedResult = table.Execute(retrieveOperation);

    //    if(retrievedResult.Result != null)
    //    {
    //        return (T)retrievedResult.Result;
    //    }
    //    else
    //    {
    //        return default(T);
    //    }
    //}

    /// <summary>
    /// Deletes a cloud object, typically requires the original object, but will work without it.
    /// </summary>
    public static void Delete(this CloudTable table, TableEntity item)
    {
        if (item.ETag == null)
        {
            item.ETag = "*";
        }

        table.Execute(TableOperation.Delete(item));
    }

    public static void Delete(this CloudTable table, string partitionKey, string rowKey)
    {
        var e = new TableEntity() { PartitionKey = partitionKey, RowKey = rowKey, ETag = "*" };
        table.Delete(e);
    }

    /// <summary>
    /// Uses the partition key / row key to update the target item. If it doesn't already exist, it
    /// will be created! (part of the service, no extra charge)
    /// </summary>
    public static void ForceUpdate(this CloudTable Table, ITableEntity item)
    {
        TableOperation update = TableOperation.InsertOrReplace(item);
        Table.Execute(update);
    }

    public static IEnumerable<DynamicTableEntity> SelectAll<T>(this CloudTable table, String partitionName = "")
    {
        TableQuery<DynamicTableEntity> query;

        if (string.IsNullOrEmpty(partitionName))
        {
            query = new TableQuery<DynamicTableEntity>();
        }
        else
        {
            query = new TableQuery<DynamicTableEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, partitionName));
        }

        return table.ExecuteQuery(query);
    }

    /// <summary>
    /// Shorthand for inserting the object ... ? Must have partitionkey, rowkey, etc.
    /// </summary>
    public static void Write<T>(this CloudTable table, T o)
    {
    }
}