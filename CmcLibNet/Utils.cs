using System;
using System.Collections.Generic;
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
        /// Encloses a string with double-quotes.
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
        /// <param name="arg">object.</param>
        /// <returns>string array.</returns>
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
        /// Returns first enum value that matches the value of the property of an attribute to an enum.
        /// </summary>
        /// <typeparam name="TEnum">Enum to inspect.</typeparam>
        /// <typeparam name="TAttr">Attribute to inspect.</typeparam>
        /// <param name="attrPropertyName">Property of attribute to search in.</param>
        /// <param name="propertyValue">Property value to search for.</param>
        /// <returns>Enum value.</returns>
        /// <exception cref="ArgumentException"></exception>
        internal static TEnum EnumFromAttributeValue<TEnum, TAttr>(string attrPropertyName, object propertyValue)
            where TEnum : struct
            where TAttr : Attribute
        {
            Type enumType = typeof(TEnum);
            Type attrType = typeof(TAttr);

            if (!enumType.IsEnum)
            {
                throw new ArgumentException($"{enumType.FullName} is not an enum.");
            }
            var propertyInfoAttributePropertyName = attrType.GetProperty(attrPropertyName);
            if (propertyInfoAttributePropertyName is null)
            {
                throw new ArgumentException($"Attribute with property {attrPropertyName} not found in enum {enumType.FullName}.");
            }

            foreach (var field in enumType.GetFields())
            {
                var attribute = Attribute.GetCustomAttribute(field, attrType);
                if (attribute == null)
                {
                    continue;
                }

                object attributePropertyValue = propertyInfoAttributePropertyName.GetValue(attribute);

                if (attributePropertyValue == null)
                {
                    continue;
                }

                if (attributePropertyValue.Equals(propertyValue))
                {
                    return (TEnum)field.GetValue(null);
                }
            }
            throw new ArgumentException($"No {attrType.FullName} attribute with property {attrPropertyName} in enum '{enumType.FullName} has value {propertyValue}.");
        }

        /// <summary>
        /// Replace all non-letter and non-digit characters.
        /// </summary>
        /// <param name="str">Input string</param>
        /// <param name="replacement">Replacement string.</param>
        /// <returns></returns>
        internal static string EscapeString(string str, string replacement)
        {
            if (string.IsNullOrEmpty(str)) { return str; }
            string pattern = @"[^.\d\w]";
            // returns a string that contains only alfanumeric characters, everything else replaced
            return Regex.Replace(str, pattern, replacement);
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