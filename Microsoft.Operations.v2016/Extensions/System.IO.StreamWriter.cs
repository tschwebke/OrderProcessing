using System;
using System.IO;

namespace Microsoft.Operations
{
    public static class ExtendStreamWriter
    {
        /// <summary>
        /// Used for Logfile writing, automatically appends a datetime stamp to the front of the line
        /// (so you don't have to do it in the code).
        /// </summary>
        /// <param name="inputText">The text you want to commit to the file</param>
        public static void WriteEntry(this StreamWriter textFile, string inputText)
        {
            textFile.WriteLine(string.Format("{0:yyyyMMdd.HHmmss.fff} {1}", DateTime.Now, inputText));
        }

        /// <summary>
        /// Used for Logfile writing, automatically appends a datetime stamp to the front of the
        /// line. Can use arguments for formatting if desired.
        /// </summary>
        /// <param name="inputText">
        /// The text you want to commit to the file, which may have items like '{0}'
        /// </param>
        public static void WriteEntry(this StreamWriter textFile, string inputText, params object[] args)
        {
            inputText = inputText.Insert(0, string.Format("{0:yyyyMMdd.HHmmss.fff} ", DateTime.Now));
            textFile.WriteLine(inputText, args);
        }
    }
}