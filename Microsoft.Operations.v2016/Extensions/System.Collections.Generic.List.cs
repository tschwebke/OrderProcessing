using System.Collections.Generic;
using System.Linq;

public static partial class ExtensionMethods
{
    /// <summary>
    /// h/t: http://stackoverflow.com/questions/47752/remove-duplicates-from-a-listt-in-c-sharp
    /// </summary>
    public static List<T> Deduplicate<T>(this List<T> listToDeduplicate)
    {
        return listToDeduplicate.Distinct().ToList();
    }
}