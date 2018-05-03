using Microsoft.WindowsAzure.Storage.Queue;

public static class CloudQueueExtensions
{
    /// <summary>
    /// Returns the type requested, and then automatically deletes it from the queue. Use this only
    /// if you trust your code to hang onto the message for as long as it needs it. Bypasses the
    /// silly 30-second invisibility, guarantees that no other clients will read the same message.
    /// </summary>
    public static T Dequeue<T>(this CloudQueue q)
    {
        CloudQueueMessage mymessage = q.GetMessage(); // grab the first message waiting here
        var myobject = mymessage.Deserialize<T>(); // convert it back into our expected type
        q.DeleteMessage(mymessage); // remove the original message
        return myobject; // and give the requestor their custom object.
    }

    /// <summary>
    /// Use this signature for any polymorphic objects !
    /// </summary>
    public static object Dequeue(this CloudQueue q)
    {
        CloudQueueMessage mymessage = q.GetMessage(); // grab the first message waiting here
        var myobject = mymessage.Deserialize(); // convert it back into our expected type
        q.DeleteMessage(mymessage); // remove the original message
        return myobject; // and give the requestor their custom object.
    }
}