using System.Collections;
using System.Text;

public static partial class ExtensionMethods
{
    /// <summary>
    /// Credit: http://stackoverflow.com/users/404854/kbrimington
    /// </summary>
    public static string Flatten(this IEnumerable elems, string separator)
    {
        if (elems == null)
        {
            return null;
        }

        StringBuilder sb = new StringBuilder();
        foreach (object elem in elems)
        {
            if (sb.Length > 0)
            {
                sb.Append(separator);
            }

            sb.Append(elem);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Returns value used for TFS MultiSelect control (Codeplex version). [Value1];[Value2]
    /// </summary>
    public static string Flatten(this IEnumerable elems, string left, string right, string separator)
    {
        if (elems == null)
        {
            return null;
        }

        StringBuilder sb = new StringBuilder();
        foreach (object elem in elems)
        {
            if (sb.Length > 0)
            {
                sb.Append(separator);
            }

            sb.Append(left + elem.ToString() + right);
        }

        return sb.ToString();
    }
}