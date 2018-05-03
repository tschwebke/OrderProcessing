using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Microsoft.Operations
{
    public static partial class ExtensionMethods
    {
        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string After(this string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }

        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string Before(this string value, string a)
        {
            int posA = value.IndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }

        /// <summary>
        /// (Alternate Method) Get string value between [first] a and [last] b.
        /// </summary>
        public static string Between(this string value, string a, string b)
        {
            int posA = value.IndexOf(a);
            int posB = value.LastIndexOf(b);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }

        /// <summary>
        /// Turns the enumerator into a string value and uses it directly as an extension of
        /// 'Contains'. Makes for easier shortFhand when working between enumerators and strings.
        /// </summary>
        public static bool Contains(this string text, Enum enumeratorToTest)
        {
            return text.Contains(enumeratorToTest.ToString());
        }

        /// <summary>
        /// A way to get around the restriction of case-sensitive nature, ref: https://connect.microsoft.com/VisualStudio/feedback/details/435324/the-string-contains-method-should-include-a-signature-accepting-a-systen-stringcomparison-value
        /// </summary>
        /// <param name="str"></param>
        /// <param name="value"></param>
        /// <param name="comparisonType"></param>
        /// <returns></returns>
        public static bool Contains(this string str, string value, StringComparison comparisonType)
        {
            return str.IndexOf(value, comparisonType) >= 0;
        }

        /// <authors>Isaac Schlueter (http://www.linkedin.com/pub/dir/isaac/schlueter)</authors>
        public static bool EqualsIgnoreCase(this string left, string right)
        {
            return string.Compare(left, right, true) == 0;
        }

        /// <summary>
        /// Returns collection of email addresses which might be found in the string, using Regular
        /// Expressions for well-formed emails only. A blank string will return an empty List.
        /// </summary>
        public static List<string> ExtractEmailAddresses(this string candidateText)
        {
            List<string> emails = new List<string>();

            if (!string.IsNullOrEmpty(candidateText))
            {
                Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
                MatchCollection emailMatches = emailRegex.Matches(candidateText);

                foreach (Match emailMatch in emailMatches)
                {
                    emails.Add(emailMatch.Value);
                }
            }

            return emails;
        }

        /// <authors>Isaac Schlueter (http://www.linkedin.com/pub/dir/isaac/schlueter)</authors>
        public static string Fill(this string format, params object[] args)
        {
            return string.Format(format, args);
        }

        /// <summary>
        /// Generic method to find the article 'a' in a sentence before a vowel sound and replace it
        /// with article 'an'
        /// </summary>
        public static string FixedArticleBeforeVowelSound(this string textSentense)
        {
            List<string> sList = textSentense.Split(' ').ToList();

            for (int i = 0; i < sList.Count; i++)
            {
                if (sList[i] == "a" && Regex.IsMatch(sList[i + 1], "^[aeiou].*", RegexOptions.IgnoreCase))
                    sList[i] = "an";
            }

            string result = string.Empty;
            sList.ForEach(w => result += string.Format("{0} ", w));
            return result;
        }

        public static string GetTextAfter(this string input, string firstOccurenceText)
        {
            int startPos = input.IndexOf(firstOccurenceText) + firstOccurenceText.Length;
            int finishPos = input.Length;
            int length = finishPos - startPos;
            string textBetweenMarkers = string.Empty;

            if (startPos > 0)
            {
                return input.Substring(startPos, length);
            }
            else
            {
                return input;
            }
        }

        /// <summary>
        /// Returns empty if no valid match. IS CASE SENSITIVE!
        /// </summary>
        /// <param name="input"></param>
        /// <param name="firstOccurenceText"></param>
        /// <param name="lastOccurenceText"></param>
        public static string GetTextBetween(this string input, string firstOccurenceText, string lastOccurenceText)
        {
            int startPos = input.IndexOf(firstOccurenceText) + firstOccurenceText.Length;
            int finishPos = input.LastIndexOf(lastOccurenceText);
            int length = finishPos - startPos;
            string textBetweenMarkers = string.Empty;

            if (startPos > 0 && finishPos > 0)
            {
                textBetweenMarkers = input.Substring(startPos, length);
            }

            return textBetweenMarkers;
        }

        /// <summary>
        /// Determines whether the string contained whitespace or not (as its only content). This is
        /// useful shorthand for combining with some tests in the case where the string is expected
        /// to contain whitespace but nothing else. For example, if the string value is " " (i.e.
        /// contains spaces) then this method returns a more correct result:
        /// - IsNullOrEmpty = false
        /// - IsBlank = true
        /// NOTE: You can achieve the same in your own code by using the Trim() method
        /// </summary>
        public static bool IsBlank(this string s)
        {
            return string.IsNullOrEmpty(s.Trim());
        }

        /// <summary>
        /// Indicates whether the specified String object is null or an Empty string. IMPORTANT NOTE:
        /// this is a SHORTHAND wrapper for IsNullorEmpty(string).
        /// </summary>
        public static bool IsEmpty(this string s)
        {
            return string.IsNullOrEmpty(s);
        }

        /// <summary>
        /// Specifies whether the input is numeric in nature (i.e. could it be parsed into a number?)
        /// </summary>
        public static bool IsNumeric(this string value)
        {
            double d;
            return double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture.NumberFormat, out d);
            // bool isNumeric = double.TryParse(value, out d);
        }

        /// <summary>
        /// Used to check whether the email format is valid or not, using basic pattern matching from
        /// regular expressions.
        /// NOTE: Does not detect the newer (non-standard) top level domains like .asia
        /// </summary>
        /// <remarks>
        /// The exact regular expression validation is a matter of debate (see various internet
        /// forums) However, this is one of the more complete definitions that we found. The
        /// expression has been condensed to just one line for the sake of brevity.
        /// </remarks>
        /// <returns>Returns true if it has valid format (using 'IsMatch' method)</returns>
        public static bool IsValidEmailAddress(this string email)
        {
            if (string.IsNullOrEmpty(email))
            {
                return false;
            }
            else
            {
                return new Regex(@"^(([^<>()[\]\\.,;:\s@\""]+(\.[^<>()[\]\\.,;:\s@\""]+)*)|(\"".+\""))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$").IsMatch(email);
            }
        }

        /// <summary>
        /// Checks if url is valid. from http://www.osix.net/modules/article/?id=586 and changed to
        /// match http://localhost The complete (not only http) url regex can be found at http://internet.ls-la.net/folklore/url-regexpr.html
        /// </summary>
        /// <remarks>
        /// This implementation describes the parts which are used for validation, which is helpful
        /// for interpreting the logic being applied.
        /// </remarks>
        /// <authors>Tomas Kubes (tomas_kubes@seznam.cz)</authors>
        public static bool IsValidUrl(this string url)
        {
            string strRegex = "^(https?://)"
            + "?(([0-9a-z_!~*'().&=+$%-]+: )?[0-9a-z_!~*'().&=+$%-]+@)?" //user@
            + @"(([0-9]{1,3}\.){3}[0-9]{1,3}" // IP- 199.194.52.184
            + "|" // allows either IP or domain
            + @"([0-9a-z_!~*'()-]+\.)*" // tertiary domain(s)- www.
            + @"([0-9a-z][0-9a-z-]{0,61})?[0-9a-z]" // second level domain
            + @"(\.[a-z]{2,6})?)" // first level domain- .com or .museum is optional
            + "(:[0-9]{1,5})?" // port number- :80
            + "((/?)|" // a slash isn't required if there is no file name
            + "(/[0-9a-z_!~*'().;?:@&=+$,%#-]+)+/?)$";
            return new Regex(strRegex).IsMatch(url);
        }

        public static bool IsVowel(this string value)
        {
            // TODO: This ... Needs to be better. oh well.

            if (value.ToLower() == "a") return true;
            if (value.ToLower() == "e") return true;
            if (value.ToLower() == "i") return true;
            if (value.ToLower() == "o") return true;
            if (value.ToLower() == "u") return true;
            return false;
        }

        /// <summary>
        /// Using the reference string, and the nominated split value, automatically grabs the last
        /// occurrence of the value, based on the split. This is a handy shortcut to get that last value.
        /// </summary>
        /// <param name="characterSplitSequence">
        /// Single or multiple character which will be used with the 'split' routine.
        /// </param>
        /// <returns>
        /// The potential split item, assuming it exists, but if not then the original string.
        /// </returns>
        public static string LastValueOfArrayBasedSplit(this string originalText, string characterSplitSequence)
        {
            string[] parts = originalText.Split(characterSplitSequence.ToCharArray());

            if (parts.Length > 0)
            {
                return parts[parts.Length - 1];
            }
            else
            {
                return originalText;
            }
        }

        /// <summary>
        /// A very basic 'trim' function which truncates the given string using the given number of
        /// characters, starting from the left position.
        /// </summary>
        public static string Left(this string s, int len)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            else if (s.Length <= len)
                return s;
            else
                return s.Substring(0, len);
        }

        /// <summary>
        /// Forces a truncation on a given string, at a given length.
        /// TODO: This should be merged with 'Truncate', which is a better extension
        /// </summary>
        public static string MaxLength(this string candidate, int characterLimit)
        {
            if (candidate.Length < characterLimit) return candidate;
            else return candidate.Substring(0, characterLimit);
        }

        /// <authors>Isaac Schlueter (http://www.linkedin.com/pub/dir/isaac/schlueter)</authors>
        public static string RegexReplace(this string input, string pattern, string replacement)
        {
            return Regex.Replace(input, pattern, replacement);
        }

        /// <authors>Isaac Schlueter (http://www.linkedin.com/pub/dir/isaac/schlueter)</authors>
        public static string RegexReplace(this string input, string pattern, string replacement, RegexOptions options)
        {
            return Regex.Replace(input, pattern, replacement, options);
        }

        public static string RemoveInvalidFileNameCharacters(this string filename)
        {
            return string.Join(string.Empty, filename.Split(Path.GetInvalidFileNameChars()));
        }

        /// <summary>
        /// Special whitespace elimination
        /// </summary>
        public static string RemoveNonAlphanumeric(this string textSentence)
        {
            return Regex.Replace(textSentence, @"[\W]", "");
        }

        /// <authors>Isaac Schlueter (http://www.linkedin.com/pub/dir/isaac/schlueter)</authors>
        public static string RemoveRange(this string input, int startIndex, int endIndex)
        {
            return input.Remove(startIndex, endIndex - startIndex);
        }

        public static string RemoveWhitespace(this string input)
        {
            return new string(input.ToCharArray()
                .Where(c => !char.IsWhiteSpace(c))
                .ToArray());
        }

        /// <summary>
        /// Generic method to find the article 'a' in a sentence before a vowel sound and replace it
        /// with article 'an'
        /// </summary>
        /// <param name="imageID">e.g. 992e9bbf-79b3-46db-8cf2-7e0f1bec868d</param>
        public static string ReplaceImageSource(this string originalText, string imageID, string replacementSource)
        {
            if (originalText.Contains(imageID))
            {
                int slicePoint = originalText.IndexOf(imageID, 0, StringComparison.OrdinalIgnoreCase);

                string part1 = originalText.Substring(0, slicePoint);
                string part2 = originalText.Substring(slicePoint, 300);

                int slicePoint1 = part1.LastIndexOf("\"", StringComparison.OrdinalIgnoreCase);
                // int slicePoint2 = part2.IndexOf("\"/>") + part1.Length;
                int slicePoint2 = (part1.Length - slicePoint1) + imageID.Length;

                string match = originalText.Substring(slicePoint1 + 1, slicePoint2 - 1);

                string newText = originalText.Replace(match, replacementSource);

                return newText;
            }
            else
            {
                return originalText;
            }
        }

        /// <summary>
        /// Strips a string of everything except for 0-9
        /// </summary>
        public static string ToDigitsOnly(this string input)
        {
            return new string(input.Where(char.IsDigit).ToArray());
        }

        /// <summary>
        /// Flatten a list of string, using a delimiter
        /// </summary>
        public static string ToString(this List<String> collection, string delimiter)
        {
            StringBuilder sb = new StringBuilder();
            foreach (String s in collection)
            {
                sb.Append(s + delimiter);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Shortcut to the TextInfo class, requires accessing the current thread .. Use sparingly.
        /// </summary>
        public static string ToTitleCase(this string s)
        {
            CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;
            TextInfo textInfo = cultureInfo.TextInfo;
            return textInfo.ToTitleCase(s);
        }

        /// <summary>
        /// Truncates a string so that it will fit the specified length, but allows placement of text
        /// at the end.
        /// </summary>
        /// <param name="len">Number of characters which represents the ceiling of what you want</param>
        /// <param name="appendix">value which will be placed at the end of the truncation</param>
        public static string Truncate(this string s, int lenthMaximumDesired, string appendix = " ...")
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            else if (s.Length <= lenthMaximumDesired)
                return s;
            else
                return s.Substring(0, (lenthMaximumDesired - appendix.Length)) + appendix;
        }
    }
}