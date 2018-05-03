using System.Drawing;
using System.IO;

namespace Microsoft.Operations
{
    public static class ExtendImage
    {
        static public byte[] ImageToByteArray(this Image imageIn, System.Drawing.Imaging.ImageFormat outputType)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, outputType);
            return ms.ToArray();
        }
    }
}