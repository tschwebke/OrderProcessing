using System.Xml;

namespace Microsoft.Operations
{
    static public partial class ExtendXmlDocument
    {
        /// <summary>
        /// Shortcut which allows you to specify the loading of an xml file which is located at the
        /// root. Saves you the hassle of having to do the complicated code of finding the actual
        /// location on disk.
        /// TODO: XElement also has a load method, and may be more efficient.
        /// TODO: This needs some cleaning up, very sloppy (by me)
        /// </summary>
        /// <param name="doc">The document object you need to fill, which you don't mind overwriting.</param>
        /// <param name="filenameSittingInApplicationRootDirectory">
        /// name of the .config file, which (MUST) be located in the root directory.
        /// </param>
        public static XmlDocument LoadFromRootDirectory(this XmlDocument doc, string filenameSittingInApplicationRootDirectory)
        {
            XmlTextReader reader = new XmlTextReader(FileSystem.ExecutingFolderFile(filenameSittingInApplicationRootDirectory));

            reader.Read();
            doc.Load(reader);
            reader.Close();
            return doc;
        }
    }
}