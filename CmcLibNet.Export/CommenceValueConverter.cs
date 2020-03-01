using System;
using System.Globalization;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Holds the data conversion routines needed to throw Commence data about
    /// </summary>
    internal static class CommenceValueConverter
    {
        /// <summary>
        /// When a field is set to 'Display as currency' and canonical data are requested,
        /// the currency will be '$' except when data comes from related categories,
        /// in which case it can be any currency symbol.
        /// </summary>
        internal static readonly string CanonicalCurrencySymbol = "$";

        /// <summary>
        /// Converts regular Commence value to its canonical representation.
        /// </summary>
        /// <remarks>Commence can return canonical values, but only from direct fields.
        /// This will convert all values to their canonical representation.
        /// </remarks>
        /// <param name="regularCommenceValue">regular Commence value.</param>
        /// <param name="fieldType">Commence fieldtype.</param>
        /// <returns>string.</returns>
        internal static string ToCanonical(string regularCommenceValue, Vovin.CmcLibNet.Database.CommenceFieldType fieldType)
        {
            // TODO: should we include a check if a canonical value was passed in?

            // Assumes we get a value from Commence, *not* some random string value.
            // The Commence value will be formatted according to system regional settings,
            // because that is the format Commence will display and return them.
            // We take advantage of that in TryParse to the proper .NET data type,
            // then we deconstruct that data type to return a string in canonical format.
            if (string.IsNullOrEmpty(regularCommenceValue)) { return regularCommenceValue; }

            string retval = regularCommenceValue;
            bool isCurrency = false;
            string curSymbol = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;

            switch (fieldType)
            {
                case CommenceFieldType.Sequence:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Number: // should return xxxx.yy
                    // if Commence was set to "Display as currency" value will be preceded by the system currency symbol
                    // This setting cannot be gotten from Commence, we have to resolve at runtime.
                    string numPart = regularCommenceValue;

                    if (numPart.StartsWith(curSymbol))
                    {
                        numPart = numPart.Remove(0, curSymbol.Length).Trim();
                        isCurrency = true;
                    }
                    // try to cast
                    Decimal number;
                    if (Decimal.TryParse(numPart, out number))
                    {
                        long intPart;
                        intPart = (long)Math.Truncate(number); // may fail if number is bigger than long.
                        // d may be negative
                        int decPart;
                        decPart = (int)(((Math.Abs(number) - Math.Abs(intPart)) * 100));
                        if (isCurrency)
                        {
                            retval = CanonicalCurrencySymbol + intPart.ToString() + '.' + decPart.ToString("D2");
                        }
                        else
                        {
                            retval = intPart.ToString() + '.' + decPart.ToString("D2");
                        }
                    }
                    break;
                case CommenceFieldType.Date: // should return yyyymmdd
                    DateTime date;
                    if (DateTime.TryParse(regularCommenceValue, out date))
                    {
                        retval = date.Year.ToString("D4") + date.Month.ToString("D2") + date.Day.ToString("D2");
                    }
                    break;
                case CommenceFieldType.Time: // should return hh:mm
                    DateTime time;
                    if (DateTime.TryParse(regularCommenceValue, out time))
                    {
                        retval = time.Hour.ToString("D2") + ':' + time.Minute.ToString("D2");
                    }
                    break;
                case CommenceFieldType.Checkbox: // should return 'true' or 'false' (lowercase)
                    // direct checkbox fields are returned as 'TRUE' or 'FALSE'
                    if (regularCommenceValue.Equals("TRUE") || regularCommenceValue.Equals("FALSE"))
                    {
                        retval = regularCommenceValue.ToLower(); // all OK
                    }
                    else
                    {
                        // related checkbox fields return 'Yes' or 'No'
                        retval = (regularCommenceValue.ToLower() == "yes").ToString().ToLower(); // Only returns true if 'yes'. Potential issue here?
                    }
                    break;
            }
            return retval;
        }
        /// <summary>
        /// Converts canonical value to string representation ISO8601-compliant value.
        /// </summary>
        /// <param name="canonicalValue">canonical Commence value.</param>
        /// <param name="fieldType">Commence fieldtype.</param>
        /// <returns></returns>
        internal static string ToIso8601(string canonicalValue, CommenceFieldType fieldType)
        {
            // assumes Commence data in value are already in canonical format!
            if (string.IsNullOrEmpty(canonicalValue)) { return canonicalValue; }

            string retval = canonicalValue;
            string currencySymbol = CanonicalCurrencySymbol; // when canonical, Commence always uses '$' regardless of regional settings

            switch (fieldType)
            {
                // Commence may return a currency symbol if the field is defined to display as currency.
                // There is no way of requesting that setting, we have to resolve it at runtime.
                // When in canonical mode, it is always a dollar sign.
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                    if (canonicalValue.StartsWith(currencySymbol))
                    {
                        retval = canonicalValue.Remove(0, currencySymbol.Length);
                    }
                    break;
                case CommenceFieldType.Date: // expects "yyyymmdd"
                    retval = canonicalValue.Substring(0, 4) + "-" + canonicalValue.Substring(4, 2) + "-" + canonicalValue.Substring(6, 2);
                    break;
                case CommenceFieldType.Time: // expects "hh:mm"
                    string[] s = canonicalValue.Split(':');
                    retval = s[0] + ":" + s[1] + ":00";
                    break;
                case CommenceFieldType.Checkbox: // expects "TRUE" or "FALSE" (case-insensitive)
                    retval = canonicalValue.ToLower();
                    break;
            }
            return retval;
        }

        /// <summary>
        /// Converts canonical value to object that ADO.NET understands.
        /// </summary>
        /// <param name="canonicalValue">Canonical Commence value.</param>
        /// <param name="fieldType">Commence fieldtype.</param>
        /// <returns>Object of type that Ado understands, DBNull.Value on empty.</returns>
        internal static object ToAdoNet(string canonicalValue, CommenceFieldType fieldType)
        {
            if (string.IsNullOrEmpty(canonicalValue)) { return DBNull.Value; }

            object retval = DBNull.Value;

            switch (fieldType)
            {
                case CommenceFieldType.Date:
                    int yearPart = Convert.ToInt32(canonicalValue.Substring(0, 4));
                    int monthPart = Convert.ToInt32(canonicalValue.Substring(4, 2));
                    int dayPart = Convert.ToInt32(canonicalValue.Substring(6, 2));
                    retval = new DateTime(yearPart, monthPart, dayPart);
                    break;
                case CommenceFieldType.Time:
                    int hourPart = Convert.ToInt32(canonicalValue.Substring(0, 2));
                    int minutePart = Convert.ToInt32(canonicalValue.Substring(3, 2));
                    retval = new DateTime(1971, 12, 23, hourPart, minutePart, 0);
                    break;
                case CommenceFieldType.Sequence:
                    retval = Convert.ToDecimal(canonicalValue);
                    break;
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                    retval = RemoveCurrencySymbol(canonicalValue);
                    retval = Convert.ToDecimal(retval);
                    break;
                default:
                    retval = canonicalValue;
                    break;
            }
            return retval;
        }
        /// <summary>
        /// Removes leading currency symbol from string.
        /// </summary>
        /// <param name="value">string.</param>
        /// <returns></returns>
        internal static string RemoveCurrencySymbol(string value)
        {
            string negativeSymbol = string.Empty;
            if (value.StartsWith("-")) // AVB 2019-05-24
            {
                value = value.Right(value.Length - 1);
                negativeSymbol = "-";
            }
            if (value.StartsWith(CanonicalCurrencySymbol))
            {
                string ret = value;
                ret = value.Substring(1).Trim();
                return ret;
            }
            if (value.StartsWith(CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol))
            {
                return negativeSymbol + value.Remove(0, CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol.Length).Trim();
            }
            else
            {
                return negativeSymbol + value;
            }
        }
    }

}
