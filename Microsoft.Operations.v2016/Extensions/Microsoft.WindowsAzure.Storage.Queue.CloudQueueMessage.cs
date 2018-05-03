using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using System;
using System.Text;

/// <summary>
/// var myobject = new MyObject(); _queue.AddMessage(
/// CloudQueueMessageExtensions.Serialize(myobject)); var myobject = _queue.GetMessage().Deserialize
/// <![CDATA[<MyObject>()]]> ;
/// </summary>
public static class CloudQueueMessageExtensions
{
    /// <summary>
    /// Polymorphic
    /// </summary>
    public static object Deserialize(this CloudQueueMessage m)
    {
        int indexOf = m.AsString.IndexOf(':');

        string typeName = m.AsString.Substring(0, indexOf);
        string json = m.AsString.Substring(indexOf + 1);

        var settings = new JsonSerializerSettings();
        settings.TypeNameHandling = TypeNameHandling.Objects;

        return JsonConvert.DeserializeObject<dynamic>(json, settings);
    }

    /// <summary>
    /// Explicit type ...
    /// </summary>
    public static T Deserialize<T>(this CloudQueueMessage m)
    {
        int indexOf = m.AsString.IndexOf(':');

        if (indexOf <= 0)
            throw new Exception(string.Format("Cannot deserialize into object of type {0}",
                typeof(T).FullName));

        string typeName = m.AsString.Substring(0, indexOf);
        string json = m.AsString.Substring(indexOf + 1);

        if (typeName != typeof(T).FullName)
        {
            throw new Exception(string.Format("Cannot deserialize object of type {0} into one of type {1}",
                typeName,
                typeof(T).FullName));
        }

        return JsonConvert.DeserializeObject<T>(json);
    }

    public static CloudQueueMessage Serialize(Object o)
    {
        var stringBuilder = new StringBuilder();
        var settings = new JsonSerializerSettings();
        settings.TypeNameHandling = TypeNameHandling.Objects;

        stringBuilder.Append(o.GetType().FullName);
        stringBuilder.Append(':');
        stringBuilder.Append(JsonConvert.SerializeObject(o, Formatting.Indented, settings));
        return new CloudQueueMessage(stringBuilder.ToString());
    }
}