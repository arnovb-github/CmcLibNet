using System;
using System.Reflection;
using System.ComponentModel;
using System.Linq;
using Vovin.CmcLibNet.Database;

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
        /// Extension method for enums that returns description attributes for enum value.
        /// </summary>
        /// <param name="value">Enum value.</param>
        /// <returns>String representation of enum member, or its description (if defined).</returns>
        public static string GetEnumDescription(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
        /// <summary>
        /// Return enum value from description.
        /// </summary>
        /// <typeparam name="T">Enum.</typeparam>
        /// <param name="description">Description of enum value to search for.</param>
        /// <returns>Enum value matching description.</returns>
        /// <exception cref="InvalidOperationException">No enum provided.</exception>
        public static T GetValueFromDescription<T>(string description)
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
        internal static Type GetSystemTypeForCommenceField(CommenceFieldType fieldType)
        {
            // String type DataTable columns can be specified to have a certain length,
            // but by default they take any length,
            // so there is no need to request the length from Commence.
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                    return typeof(Decimal);
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return typeof(DateTime);
                case CommenceFieldType.Sequence:
                    return typeof(UInt32);
                case CommenceFieldType.Checkbox:
                    return typeof(Boolean);
                default:
                    return typeof(String);
            }
        }
    }
}
