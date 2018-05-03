using System.Collections.Generic;
using System.Xml.Linq;

public static class ExtendDictionary
{
    /// <summary>
    /// Very simple routine to output contents as readable HTML (intended for Outlook HTML) Used
    /// mostly for trouble-shooting.
    /// </summary>
    //public static string AsHtmlTable(this Dictionary<string, string> coll)
    //{
    //    StringBuilder sb = new StringBuilder();

    // if (items.Count > 0) { sb.Append("<table>");

    // foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field info in result) {
    // Console.WriteLine(wi.Id + info.Status); }

    // sb.Append("</table>");

    // }

    //    return sb.ToString();
    //}

    /// <summary>
    /// This method exists sort of in the .NET 4.0 full library ... part of the 'ConcurrentDictionary'
    /// </summary>
    /// <param name="coll">The existing string dictionary that you wish to update</param>
    /// <param name="key">Standard key name, which is expected to be found in dictionary</param>
    /// <param name="value">Value which shall be stored (or updated)</param>
    public static void AddOrUpdate(this Dictionary<string, string> coll, string key, string value)
    {
        coll.AddOrUpdate(key, value, true);
    }

    /// <summary>
    /// </summary>
    /// <param name="coll"></param>
    /// <param name="key"></param>
    /// <param name="value"></param>
    /// <param name="forceOverwrite">
    /// decide what to do with the existing value if it is present (leave intact = false) or
    /// (overwrite = true).
    /// </param>
    public static void AddOrUpdate(this Dictionary<string, string> coll, string key, string value, bool forceOverwrite)
    {
        if (!coll.ContainsKey(key))
        {
            coll.Add(key, value); // add
        }
        else
        {
            if (forceOverwrite)
            {
                coll[key] = value; // overwrite
            }
        }
    }

    /// <summary>
    /// Takes an additional collection and updates the original one. This version overwrites the
    /// original with the additional content, if the key is the same.
    /// </summary>
    public static void MergeWith(this Dictionary<string, string> original, Dictionary<string, string> additional)
    {
        foreach (string s in additional.Keys)
        {
            original.AddOrUpdate(s, additional[s]);
        }
    }

    /// <summary>
    /// Takes a simple string dictionary object and turns the VALUES (not the keys) into a list of Strings.
    /// </summary>
    public static List<string> ToStringList(this Dictionary<string, string> coll)
    {
        List<string> listFromDictionary = new List<string>();
        foreach (KeyValuePair<string, string> entry in coll)
        {
            listFromDictionary.Add(entry.Value);
        }
        return listFromDictionary;
    }

    public static XElement ToXElement(this Dictionary<string, string> coll)
    {
        XElement root = new XElement("StringDictionary");
        foreach (KeyValuePair<string, string> entry in coll)
        {
            root.Add(new XElement(entry.Key, entry.Value));
        }
        return root;
    }
}