using System;
using System.IO;
using System.Reflection;

namespace Microsoft.Operations
{
    public static class FileSystem
    {
        /// <summary>
        /// The location of a folder which can be used by the application (normally used for Logfiles
        /// and Temporary files). Use this location in favour of any other 'special' location for
        /// files, always has permission to write here (is the officially preferred)
        /// </summary>
        public static string BaseFolder
        {
            get
            {
                return string.Format(@"{0}\Microsoft Operations", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData));
            }
        }

        /// <summary>
        /// Using reflection, the location of the folder for the executing binaries. Credits to
        /// StackOverflow (scrapped our original)
        /// </summary>
        public static string ExecutingFolder
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        /// <summary>
        /// MAKE THIS OBSOLETE
        /// </summary>
        public static string ExecutingFolderFile(string fileName)
        {
            return string.Format(@"{0}\{1}", FileSystem.ExecutingFolder, fileName);
        }

        /// <summary>
        /// For resource files where the content is known to be string, this is an easy and efficient
        /// way to extract the contents. The name of the resource/file within ... e.g. 'Microsoft.Operations.Webservices.Templates.QueryCapacityAllocations.wiq'
        /// TODO: Need a version which defaults to calling assembly (for a shorter signature)
        /// </summary>
        public static string GetEmbeddedFileContent(string fileName, string assemblyName)
        {
            // Check the contents of the cache for this item first ...
            string strContent = StringCache.Read(fileName);

            // Otherwise, nothing is there so we have to fetch a value and then cache it for next time.
            if (string.IsNullOrEmpty(strContent))
            {
                StreamReader streamReader = new StreamReader(Assembly.Load(assemblyName).GetManifestResourceStream(fileName));
                strContent = streamReader.ReadToEnd();
                streamReader.Close();

                StringCache.Write(fileName, strContent);
            }

            return strContent;
        }

        public static StreamWriter Logfile(string applicationName)
        {
            return Logfile(applicationName, string.Empty, string.Empty);
        }

        public static StreamWriter Logfile(string applicationName, string purpose)
        {
            return Logfile(applicationName, purpose, string.Empty);
        }

        /// <summary>
        /// A StreamWriter object which can be used to write logfile entries (with 24-hour rollover)
        /// Automatically determines location (ProgramData folder) and will reuse an existing day if
        /// it exists. Current code assumes no locks on the file specified. UPDATED to use very
        /// unique filename ...
        /// </summary>
        public static StreamWriter Logfile(string applicationName, string purpose, string uniqueMarker)
        {
            string logfileLocation = string.Format(@"{0}\{1}\Logfiles", FileSystem.BaseFolder, applicationName);
            // string logfileLocation = string.Format(@"{0}\{1}", FileSystem.BaseFolder,
            // applicationName); string logfileLocation = string.Format(@"{0}\{1}\{2}",
            // FileSystem.BaseFolder, applicationName, uniqueMarker); // should be separate subfolder
            string log_file_name;

            if (string.IsNullOrEmpty(uniqueMarker))
            {
                log_file_name = string.Format(@"{0}\{1}.{2:yyyyMMdd.HHmmss.fff}.txt", logfileLocation, purpose, DateTime.Now);
            }
            else // use/combine their unique reference.
            {
                log_file_name = string.Format(@"{0}\{1}.{2:yyyyMMdd.HHmmss.fff}.{3}.txt", logfileLocation, purpose, DateTime.Now, uniqueMarker);
            }

            StreamWriter logger;

            if (!Directory.Exists(logfileLocation))
            {
                Directory.CreateDirectory(logfileLocation);
            }

            if (File.Exists(log_file_name))
            {
                logger = File.AppendText(log_file_name);
            }
            else
            {
                logger = new StreamWriter(log_file_name);
            }

            logger.AutoFlush = true;
            return logger;
        }

        /// <summary>
        /// Allows calling code to specify a new working folder intside the ProgramData, and will
        /// create if not already existing.
        /// </summary>
        public static string WorkingFolder(string applicationName, string nameOfSubfolder)
        {
            string temp = string.Format(@"{0}\{1}\{2}", BaseFolder, applicationName, nameOfSubfolder);

            if (!Directory.Exists(temp))
            {
                Directory.CreateDirectory(temp);
            }

            return temp;
        }
    }
}