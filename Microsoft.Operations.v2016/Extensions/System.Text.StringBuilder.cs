using System.Text;

public static partial class ExtendStringBuilder
{
    public static StringBuilder AppendLine(this StringBuilder builder, string format, params object[] args)
    {
        builder.AppendFormat(format, args).AppendLine();
        return builder;
    }
}