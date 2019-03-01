using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Utility class for performing some common operations.
    /// </summary>
    internal static class Utils
    {
        /// <summary>
        /// Encloses a string with double-quotes. Does not check if string is already quoted.
        /// </summary>
        /// <param name="s">string.</param>
        /// <returns>double-quoted string.</returns>
        internal static string dq(string s)
        {
            return '"' + s + '"';
        }

        /// <summary>
        /// Creates string array from object.
        /// </summary>
        /// <param name="arg">object</param>
        /// <returns>string array</returns>
        internal static string[] ToStringArray(object arg)
        {
            var collection = arg as System.Collections.IEnumerable;
            if (collection != null)
            {
                return collection
                  .Cast<object>()
                  .Select(x => x.ToString())
                  .ToArray();
            }

            if (arg == null)
            {
                return new string[] { };
            }
            return new string[] { arg.ToString() };
        }

        /// <summary>
        /// Return enum value from description.
        /// </summary>
        /// <typeparam name="T">Enum.</typeparam>
        /// <param name="description">Description of enum value to search for.</param>
        /// <returns>Enum value matching description.</returns>
        /// <exception cref="InvalidOperationException">No enum provided.</exception>
        public static T GetValueFromEnumDescription<T>(string description)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();
            foreach (var field in type.GetFields())
            {
                var attribute = Attribute.GetCustomAttribute(field,
                    typeof(DescriptionAttribute)) as DescriptionAttribute;
                if (attribute != null)
                {
                    if (attribute.Description == description)
                        return (T)field.GetValue(null);
                }
                else
                {
                    if (field.Name == description)
                        return (T)field.GetValue(null);
                }
            }
            //throw new ArgumentException("Not found.", "description");
            return default(T); // returns the default value. May be hard to debug! TODO
        }

        /// <summary>
        /// Gets the System.Type for the Commence fieldtype
        /// </summary>
        /// <param name="fieldType">Commence field type</param>
        /// <returns>System.Type</returns>
        internal static Type GetTypeForCommenceField(CommenceFieldType fieldType)
        {
            // String type DataTable columns can be specified to have a certain length,
            // but by default they take any length,
            // so there is no need to request the length from Commence.
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                    return typeof(double); // TODO is this precise enough? Commence takes what? 8 significant numbers? Decimal is probably overkill
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return typeof(DateTime);
                case CommenceFieldType.Sequence:
                    return typeof(int);
                case CommenceFieldType.Checkbox:
                    return typeof(bool);
                default:
                    return typeof(string);
            }
        }

        internal static DocumentFormat.OpenXml.Spreadsheet.CellValues GetTypeForOpenXml(CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return DocumentFormat.OpenXml.Spreadsheet.CellValues.Number; // TODO is this precise enough? Commence takes what? 8 significant numbers? Decimal is probably overkill
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
                case CommenceFieldType.Checkbox:
                    return DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
                default:
                    return DocumentFormat.OpenXml.Spreadsheet.CellValues.String; // should ideally be SharedString for filesize
            }
        }

        internal static string GetOleDbTypeStringForCommenceField(CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return "double"; // TODO is this precise enough? Commence takes what? 8 significant numbers? Decimal is probably overkill
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return "datetime";
                case CommenceFieldType.Checkbox:
                    return "bit";
                case CommenceFieldType.Name:
                    return "char (50)";
                default:
                    return "memo"; // just make it big :)
            }
        }

        internal static DbType GetDbTypeForCommenceField(CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return DbType.Double; // TODO is this precise enough? Commence takes what? 8 significant numbers? Decimal is probably overkill
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return DbType.DateTime;
                case CommenceFieldType.Checkbox:
                    return DbType.Boolean;
                default:
                    return DbType.String;
            }
        }

        internal static OleDbType GetOleDbTypeForCommenceField(CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return OleDbType.Double; // TODO is this precise enough? Commence takes what? 8 significant numbers? Decimal is probably overkill
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return OleDbType.Date;
               case CommenceFieldType.Checkbox:
                    return OleDbType.Boolean;
                default:
                    return OleDbType.LongVarWChar;
            }
        }

        /// <summary>
        /// Replace all non-letter and non-digit characters.
        /// </summary>
        /// <param name="str">Input string</param>
        /// <param name="replaceInvalidCharsWith">Replacement string.</param>
        /// <returns></returns>
        internal static string EscapeString(string str, string replaceInvalidCharsWith)
        {
            if (string.IsNullOrEmpty(str)) { return str; }
            string pattern = @"[^.\d\w]"; // TODO check for multiple occurences
            // returns a string that contains only alfanumeric characters, everything else replaced
            return Regex.Replace(str, pattern, replaceInvalidCharsWith.ToString());
        }

        /// <summary>
        /// Removes all control characters (ascii &lt; 32) except TAB, CR and LF from a string.
        /// </summary>
        /// <param name="str">Input string.</param>
        /// <returns>String without control characters.</returns>
        internal static string RemoveControlCharacters(string str)
        {
            //// Linq implementation is more elegant but slightly slower
            //return new string(
            //    str.Select(c => (int)c)
            //    .Where(i => i >= 32 || i == 9 || i == 10 || i == 13)
            //    .Select(i => (char)i)
            //    .ToArray());

            // char arrays are fast
            char[] c = str.ToCharArray();
            StringBuilder sb = new StringBuilder();
            foreach (char x in c)
            {
                if (x > 31 || x == 9 || x == 10 || x == 13)
                {
                    sb.Append(x);
                }
            }
            return sb.ToString();
        }

        internal static string AddUniqueIdentifier(string testString, List<string> list, uint appendNumber, uint maxIterations, uint maxLength)
        {
            string retval = testString;
            uint append = appendNumber;
            // if the list items already contain a trailing number,
            // all we want to do is increase that number instead of appending one
            // so let's analyze the list items first
            // No, let's limit that to  trailing number that contains (1), (2), etc.
            // first analyze which list items have a trailing '(number)'part
            string pattern = @"\(\d+\)$";
            IEnumerable<string> existingStrings = GetStringsWithMatchingRegexPattern(list, pattern);
            // get the one with the highest numer
            if (existingStrings.Count() > 0)
            {
                // extract the numbers
                List<uint> numbers = new List<uint>();
                foreach (string s in existingStrings)
                {
                    Regex r = new Regex(@"\d+\)$");
                    Match match = r.Match(s);
                    if (match.Success)
                    {
                        numbers.Add(Convert.ToUInt32(match.Value.Left(match.Value.Length-1)));
                    }
                }
                append = numbers.Max();
            }

            // prevent eternal loop if specified
            if (maxIterations > 0)
            {
                if (append > maxIterations) { throw new Exception("Could not get a unique name for column " + testString); }
            }
            if (list.Contains(retval))
            {
                append++;
                retval = AppendBracketedString(testString,append.ToString());
                // need to check for length
                if (maxLength > 0)
                {
                    while (retval.Length > maxLength)
                    {
                        testString = testString.Left(testString.Length - 1); // TODO do we require a check to prevent Left on on empty string?
                        retval = AppendBracketedString(testString, append.ToString());
                    }
                }
                else
                {
                    retval = AppendBracketedString(testString, append.ToString());
                }
                
                return AddUniqueIdentifier(retval, list, append, maxIterations, maxLength); // recurse
            }
            else
            {
                return retval;
            }
        }

        private static string AppendBracketedString(string s, string appendString)
        {
            StringBuilder sb = new StringBuilder(s);
            sb.Append("(");
            sb.Append(appendString);
            sb.Append(")");
            return sb.ToString();
        }

        private static IEnumerable<string> GetStringsWithMatchingRegexPattern(IEnumerable<string> list, string pattern)
        {
            Regex r = new Regex(pattern);
            foreach (string s in list)
            {
                if (r.IsMatch(s))
                {
                    yield return s;
                }
            }
        }
    }
}
