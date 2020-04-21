using System;
using System.Globalization;
using Vovin.CmcLibNet.Database;

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
        private static readonly string CanonicalCurrencySymbol = "$";
        private static readonly string minusSymbol = "-";

        /// <summary>
        /// Converts regular Commence value to its canonical representation.
        /// </summary>
        /// <remarks>Commence can return canonical values, but only from direct fields.
        /// This will convert all values to their canonical representation.
        /// </remarks>
        /// <param name="regularCommenceValue">regular Commence value.</param>
        /// <param name="fieldType">Commence fieldtype.</param>
        /// <param name="removeCurrencySymbol">Omit currency symbol.</param>
        /// <returns>string.</returns>
        internal static string ToCanonical(string regularCommenceValue, 
            Vovin.CmcLibNet.Database.CommenceFieldType fieldType,
            bool removeCurrencySymbol)
        {
            // Assumes we get a value from Commence, *not* some random string value.
            // The Commence value will be formatted according to system regional settings,
            // because that is the format Commence will display and return them.
            // We take advantage of that in TryParse to the proper .NET data type,
            // then we deconstruct that data type to return a string in canonical format.
            if (string.IsNullOrEmpty(regularCommenceValue)) { return regularCommenceValue; }

            string retval = regularCommenceValue;
            bool isCurrency = false;
            string curSymbol = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;
            decimal number;

            switch (fieldType)
            {
                case CommenceFieldType.Sequence:
                    if (decimal.TryParse(retval, out number))
                    {
                        retval = Math.Truncate(number).ToString();
                    };
                    break;
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Number: // should return xxxx.yy
                    // if Commence was set to "Display as currency" value will be preceded by the system currency symbol
                    // This setting cannot be gotten from Commence, we have to resolve it at runtime.
                    string numPart = retval;
                    if (numPart.Contains(curSymbol)) { isCurrency = true; }
                    numPart = RemoveCurrencySymbol(numPart);
                    // try to cast
                    if (decimal.TryParse(numPart, out number))
                    {
                        long longPart = (long)Math.Truncate(number); // may fail if number is bigger than long.
                        int decPart;
                        decPart = (int)((Math.Abs(number) - Math.Abs(longPart)) * 100);
                        if (isCurrency && !removeCurrencySymbol)
                        {
                            retval = longPart < 0
                                ? minusSymbol + CanonicalCurrencySymbol + Math.Abs(longPart).ToString() + '.' + decPart.ToString("D2")
                                : CanonicalCurrencySymbol + longPart.ToString() + '.' + decPart.ToString("D2");
                        }
                        else
                        {
                            retval = longPart.ToString() + '.' + decPart.ToString("D2");
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
                        retval = (regularCommenceValue.ToLower() == "yes").ToString().ToLower();
                    }
                    break;
            }
            return retval;
        }

        /// <summary>
        /// Converts regular Commence value to its canonical representation.
        /// </summary>
        /// <remarks>Commence can return canonical values, but only from direct fields.
        /// This will convert all values to their canonical representation.
        /// </remarks>
        /// <param name="value">regular Commence value.</param>
        /// <param name="fieldType">Commence fieldtype.</param>
        /// <param name="settings">Settings.</param>
        /// <returns>string.</returns>
        internal static string ToCanonical(string value, 
            Vovin.CmcLibNet.Database.CommenceFieldType fieldType,
            IExportSettings settings)
        {
            // Assumes we get a value from Commence, *not* some random string value.
            // The Commence value will be formatted according to system regional settings,
            // because that is the format Commence will display and return them.
            // We take advantage of that in TryParse to the proper .NET data type,
            // then we deconstruct that data type to return a string in canonical format.
            if (string.IsNullOrEmpty(value)) { return value; }
            string retval = value;
            bool isCurrency = false;
            string curSymbol = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;
            decimal number;

            switch (fieldType)
            {
                case CommenceFieldType.Sequence:
                    if (decimal.TryParse(retval, out number))
                    {
                        retval = Math.Truncate(number).ToString();
                    };
                    break;
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Number: // should return xxxx.yy
                    // if Commence was set to "Display as currency" value will be preceded by the system currency symbol
                    // This setting cannot be gotten from Commence, we have to resolve it at runtime.
                    string numPart = retval;
                    if (numPart.Contains(curSymbol)) { isCurrency = true; }
                    numPart = RemoveCurrencySymbol(numPart);
                    // try to cast
                    if (decimal.TryParse(numPart, out number))
                    {
                        long longPart = (long)Math.Truncate(number); // may fail if number is bigger than long.
                        int decPart;
                        decPart = (int)((Math.Abs(number) - Math.Abs(longPart)) * 100);
                        if (isCurrency && !settings.RemoveCurrencySymbol)
                        {
                            retval = longPart < 0
                                ? minusSymbol + CanonicalCurrencySymbol + Math.Abs(longPart).ToString() + '.' + decPart.ToString("D2")
                                : CanonicalCurrencySymbol + longPart.ToString() + '.' + decPart.ToString("D2");
                        }
                        else
                        {
                            retval = longPart.ToString() + '.' + decPart.ToString("D2");
                        }
                    }
                    break;
                case CommenceFieldType.Date: // should return yyyymmdd
                    DateTime date;
                    if (DateTime.TryParse(value, out date))
                    {
                        if (settings.ISO8601Format)
                        {
                            retval = ToIso8601(date, fieldType);
                        }
                        else
                        {
                            retval = date.Year.ToString("D4") + date.Month.ToString("D2") + date.Day.ToString("D2");
                        }
                    }
                    break;
                case CommenceFieldType.Time: // should return hh:mm
                    DateTime time;
                    if (DateTime.TryParse(value, out time))
                    {
                        if (settings.ISO8601Format)
                        {
                            retval = ToIso8601(time, fieldType);
                        }
                        else
                        {
                            retval = time.Hour.ToString("D2") + ':' + time.Minute.ToString("D2");
                        }
                    }
                    break;
                case CommenceFieldType.Checkbox: // should return 'true' or 'false' (lowercase)
                    // direct checkbox fields are returned as 'TRUE' or 'FALSE'
                    if (value.Equals("TRUE") || value.Equals("FALSE"))
                    {
                        retval = value.ToLower(); // all OK
                    }
                    else
                    {
                        // related checkbox fields return 'Yes' or 'No'
                        retval = (value.ToLower() == "yes").ToString().ToLower();
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
        // TODO: we should decouple this from Canonical.
        // As it is now, when ISO8601 is selected, so is Canonical
        // That leads to unnecessarily complex code.
        // Also, it masks the fact that ISO8601 is only for date/time values
        // It is still not trivial though, because complex exports expect values in a format that Commence may not export
        // Also, converting them back will require some extra plumbing (see UseThids stuff)
        internal static string ToIso8601(string canonicalValue, CommenceFieldType fieldType)
        {
            string retval = canonicalValue;
            switch (fieldType)
            {
                case CommenceFieldType.Date: // expects "yyyymmdd", returns "yyyy-MM-dd"
                    retval = canonicalValue.Substring(0, 4) + '-' + canonicalValue.Substring(4, 2) + '-' + canonicalValue.Substring(6, 2);
                    break;
                case CommenceFieldType.Time: // expects "hh:mm", returns "hh:mm:ss"
                    string[] s = canonicalValue.Split(':');
                    retval = s[0] + ':' + s[1] + ":00";
                    break;
            }
            return retval;
        }

        private static string ToIso8601(DateTime datetime, CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Date:
                    return datetime.Year + "-" + datetime.Month.ToString("D2") + "-" + datetime.Day.ToString("D2");
                case CommenceFieldType.Time:
                    return datetime.Hour.ToString("D2") + ":" + datetime.Minute.ToString("D2") + ":00";
                default:
                    return string.Empty; // risky?
            }
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

            object retval;
            switch (fieldType)
            {
                case CommenceFieldType.Date:
                    canonicalValue = canonicalValue.Replace("-", string.Empty); // if ISO8601, canonical value is overridden
                    int yearPart = Convert.ToInt32(canonicalValue.Substring(0, 4));
                    int monthPart = Convert.ToInt32(canonicalValue.Substring(4, 2));
                    int dayPart = Convert.ToInt32(canonicalValue.Substring(6, 2));
                    retval = new DateTime(yearPart, monthPart, dayPart);
                    break;
                case CommenceFieldType.Time:
                    int hourPart = Convert.ToInt32(canonicalValue.Substring(0, 2));
                    int minutePart = Convert.ToInt32(canonicalValue.Substring(3, 2));
                    retval = new DateTime(1970, 1, 1, hourPart, minutePart, 0);
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
            // this method does not use the much simpler string.Replace() method
            // because the input string can have several formats
            // and we want to perforn a Trim() on the 'inner' value
            // e.g. if we get "-USD 100.99" we want to return "-100.99" not "- 100.99"
            // the space after USD may or may not be there,
            // depending on how Commence returns data
            // we could just do Replace() four times.
            string prefix = string.Empty;
            if (value.StartsWith(minusSymbol))
            {
                value = value.Remove(0, minusSymbol.Length);
                prefix = minusSymbol;
            }

            if (value.StartsWith(CanonicalCurrencySymbol))
            {
                value = value.Remove(0, CanonicalCurrencySymbol.Length).Trim();
            }

            if (value.StartsWith(CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol))
            {
                value = value.Remove(0, CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol.Length).Trim();
            }
            return prefix + value;
        }
    }

}