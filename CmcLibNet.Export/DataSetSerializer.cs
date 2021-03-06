﻿using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Exports data from a DataSet.
    /// </summary>
    internal class DataSetSerializer
    {
        DataSet _ds = null;
        IExportSettings _settings = null;
        readonly string _filename = string.Empty;

        #region Constructors
        internal DataSetSerializer(DataSet dataset, string fileName, IExportSettings settings)
        {
            _ds = dataset;
            _settings = settings;
            _filename = fileName;
        }
        #endregion

        #region Methods
        internal void Export()
        {
            switch (_settings.ExportFormat)
            {
                case ExportFormat.Xml: // never used
                    DataSetToXML(_ds);
                    break;
                case ExportFormat.Json:
                    DataSetToJson(_ds);
                    break;
                case ExportFormat.Excel:
                    DataSetToExcel(_ds, _filename);
                    break;
            }
        }

        private void DataSetToXML(DataSet ds)
        {
            // interesting fact: even with a schema included, Excel can't deal with the resulting XML!
            XmlWriterSettings xws = new XmlWriterSettings
            {
                Indent = true,
                Encoding = System.Text.Encoding.UTF8
            };
            using (System.Xml.XmlWriter xw = System.Xml.XmlWriter.Create(_filename, xws))
            {
                _ds.WriteXml(xw, XmlWriteMode.WriteSchema); // write XML schema as well.
            }
        }

        private void DataSetToJson(DataSet ds)
        {
            // see: http://stackoverflow.com/questions/21727144/convert-dataset-with-multiple-datatables-to-json
            using (StreamWriter sw = new StreamWriter(_filename))
            {
                sw.Write(JsonConvert.SerializeObject(_ds, Newtonsoft.Json.Formatting.Indented));
            }
        }

        private void DataSetToExcel(DataSet dataSet, string filePath)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                foreach (DataTable dataTable in dataSet.Tables)
                {
                    ExcelWorksheet workSheet = pck.Workbook.Worksheets.Add(dataTable.TableName);
                    workSheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                }
                pck.SaveAs(new FileInfo(filePath));
            }
        }
        #endregion
    }
}