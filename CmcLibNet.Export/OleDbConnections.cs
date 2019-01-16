using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace Vovin.CmcLibNet.Export
{
    // ACE options:
    //"Excel 12.0 Xml"; // For Excel 2007 XML (*.xlsx)  
    //"Excel 12.0"; // For Excel 2007 Binary (*.xlsb)  
    //"Excel 12.0 Macro"; // For Excel 2007 Macro-enabled (*.xlsm)  
    //"Excel 8.0"; // For Excel 97/2000/2003 (*.xls)  
    //"Excel 5.0"; // For Excel 5.0/95 (*.xls)

    internal class OleDbConnections
    {
        internal bool SheetExists(OleDbConnection cn, string sheetName)
        {
            return cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null)
                .AsEnumerable()
                .Where(row => row.Field<string>("Table_Name") == sheetName)
                .Select(row => row)
                .FirstOrDefault() != null;
        }

        /// <summary>
        /// Create a connection where first row contains column names
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="IMEX"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        [System.Diagnostics.DebuggerStepThrough()]
        internal string HeaderConnectionString(string FileName, int IMEX = 1)
        {
            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
            if (Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX={0};HDR=Yes;", IMEX));
            }
            else
            {
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", string.Format("Excel 12.0 Xml;IMEX={0};HDR=Yes;", IMEX));
            }

            Builder.DataSource = FileName;

            return Builder.ToString();
        }

        /// <summary>
        /// Create a connection where first row contains data
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="IMEX"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        [System.Diagnostics.DebuggerStepThrough()]
        internal string NoHeaderConnectionString(string FileName, int IMEX = 1)
        {
            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
            if (Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX={0};HDR=No;", IMEX));
            }
            else
            {
  
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", string.Format("Excel 12.0 Xml;IMEX={0};HDR=No;", IMEX));
            }

            Builder.DataSource = FileName;

            return Builder.ToString();
        }
    }
}
