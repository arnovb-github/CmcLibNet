using System;
using System.Data;
using System.Data.OleDb;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Extensions
{
    internal static class CommenceFieldTypeExtension
    {
        internal static string GetSQLiteTypeForCommenceField(this CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return "REAL";
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return "TEXT"; // TEXT expects an ISO8601 value. see: https://www.sqlite.org/datatype3.html
                case CommenceFieldType.Checkbox:
                    return "INTEGER";
                default:
                    return "TEXT";
            }
        }

        /// <summary>
        /// Gets the System.Type for the Commence fieldtype
        /// </summary>
        /// <param name="fieldType">Commence field type</param>
        /// <returns>System.Type</returns>
        internal static Type GetTypeForCommenceField(this CommenceFieldType fieldType)
        {
            // String type DataTable columns can be specified to have a certain length,
            // but by default they take any length,
            // so there is no need to request the length from Commence.
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence: // when a sequence is gotten canonical it will contain decimals
                    return typeof(double);
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return typeof(DateTime);
                case CommenceFieldType.Checkbox:
                    return typeof(bool);
                default:
                    return typeof(string);
            }
        }

        //internal static DocumentFormat.OpenXml.Spreadsheet.CellValues GetTypeForOpenXml(this CommenceFieldType fieldType)
        //{
        //    switch (fieldType)
        //    {
        //        case CommenceFieldType.Number:
        //        case CommenceFieldType.Calculation:
        //        case CommenceFieldType.Sequence:
        //            return DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
        //        case CommenceFieldType.Date:
        //        case CommenceFieldType.Time:
        //            return DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
        //        case CommenceFieldType.Checkbox:
        //            return DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
        //        default:
        //            return DocumentFormat.OpenXml.Spreadsheet.CellValues.String; // should ideally be SharedString for filesize
        //    }
        //}

        internal static string GetOleDbTypeStringForCommenceField(this CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return "double";
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

        internal static DbType GetDbTypeForCommenceField(this CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return DbType.Double;
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return DbType.DateTime;
                case CommenceFieldType.Checkbox:
                    return DbType.Boolean;
                default:
                    return DbType.String;
            }
        }

        internal static OleDbType GetOleDbTypeForCommenceField(this CommenceFieldType fieldType)
        {
            switch (fieldType)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Calculation:
                case CommenceFieldType.Sequence:
                    return OleDbType.Double;
                case CommenceFieldType.Date:
                case CommenceFieldType.Time:
                    return OleDbType.Date;
                case CommenceFieldType.Checkbox:
                    return OleDbType.Boolean;
                default:
                    return OleDbType.LongVarWChar;
            }
        }
    }
}
