using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
        internal static string Dq(string s)
        {
            return string.Format("\"{0}\"", s);
        }

        /// <summary>
        /// Creates string array from object.
        /// </summary>
        /// <param name="arg">object</param>
        /// <returns>string array</returns>
        internal static string[] ToStringArray(object arg)
        {
            if (arg is System.Collections.IEnumerable collection)
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
            if (!type.IsEnum) throw new InvalidOperationException("Type is not an Enum");
            foreach (var field in type.GetFields())
            {
                if (Attribute.GetCustomAttribute(field,
                    typeof(DescriptionAttribute)) is DescriptionAttribute attribute)
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
            return default(T); // returns the default value. May be hard to debug!
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
            string pattern = @"[^.\d\w]";
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

        internal static IEnumerable<string> RenameDuplicates(IList<string> list)
        {
            var q = list.GroupBy(x => x)
                .Select(g => new { Values = g, Count = g.Count() })
                .Where(w => w.Count > 1);
            foreach (var g in q)
            {
                int j = 1;
                foreach (string s in g.Values)
                {
                    list[list.IndexOf(s)] = s + j.ToString();
                    j++;
                }
            }
            return list;
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

        internal static string GetClarifiedItemName(string itemName, string clarifySeparator, string clarifyValue)
        {
            if (string.IsNullOrEmpty(itemName)) { return string.Empty; }
            if (!string.IsNullOrEmpty(clarifySeparator)) // connection specified as clarified
            {
                return itemName.PadRight(50) + clarifySeparator + clarifyValue.PadRight(40);
            }
            else // connection not specified as clarified
            {
                return itemName;
            }
        }
    }
}