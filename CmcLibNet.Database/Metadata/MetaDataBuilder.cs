using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Holds all logic for constructing schema information.
    /// </summary>
    internal class MetaDataBuilder : IDisposable
    {
        #region Fields
        private readonly ICommenceDatabase _db;
        IList<FormXmlMetaData> _formXmlFiles = null;
        private const string TEMPLATES_FOLDER = @"\tmplts\";
        #endregion

        #region Constructors
        public MetaDataBuilder(ICommenceDatabase db, MetaDataOptions options)
        {
            _db = db;
            MetaDataOptions = options;
        }
        #endregion

        #region Methods
        internal IDatabaseSchema BuildDatabaseSchema()
        {
            DatabaseSchema ds = null;
            IDBDef d = _db.GetDatabaseDefinition();
            ds = new DatabaseSchema(d)
            {
                Size = GetDatabaseSize(_db.Path)
            };
            var categoryNames = _db.GetCategoryNames();
            ds.Categories = GetCategoryMetadata(categoryNames).ToList();
            var files = this.FormXmlFiles;
            return ds;
        }
        #endregion

        #region Helper Methods
        private long GetDatabaseSize(string path)
        {
            // may throw access exception, but that should almost never be an issue.
            return Directory.GetFiles(path, "*", SearchOption.AllDirectories).Sum(t => (new FileInfo(t).Length));
        }

        private IEnumerable<CommenceCategoryMetaData> GetCategoryMetadata(IEnumerable<string> list)
        {
            foreach (string categoryName in list)
            {
                ICategoryDef d = _db.GetCategoryDefinition(categoryName);
                CommenceCategoryMetaData c = new CommenceCategoryMetaData(categoryName, d);

                var fields = _db.GetFieldNames(categoryName);
                c.Fields = GetFieldMetaData(categoryName, fields).ToList();

                c.Connections = _db.GetConnectionNames(categoryName).Cast<CommenceConnection>().ToList();

                var views = _db.GetViewNames(categoryName);
                c.Views = GetViewMetaData(views).ToList();

                var forms = _db.GetFormNames(categoryName);
                c.Forms = GetFormMetaData(categoryName, forms).ToList();

                c.Items = _db.GetItemCount(categoryName);
                yield return c;
            }
        }

        private IEnumerable<CommenceFormMetaData> GetFormMetaData(string categoryName, IEnumerable<string> forms)
        {
            foreach (string formName in forms)
            {
                var formFile = FormXmlFiles.FirstOrDefault(w => w.Category.Equals(categoryName) && w.Name.Equals(formName));
                CommenceFormMetaData f = new CommenceFormMetaData(formName)
                {
                    Category = categoryName,
                    Path = formFile?.Path
                };

                if (MetaDataOptions.IncludeFormScript)
                {
                    f.Script = GetFormScriptContent(categoryName, formName);
                }

                if (MetaDataOptions.IncludeFormXml)
                {
                    if (formFile == null) { break; }
                    using (StreamReader sr = new StreamReader(formFile.Path))
                    {
                        f.Xml = sr.ReadToEnd();
                    }
                }
                yield return f;
            }
        }

        private string GetFormScriptContent(string categoryName, string formName)
        {
            string content = "An error occurred while reading the form script.";
            string fileName = Path.GetTempFileName();
            try
            {
                if (!_db.CheckOutFormScript(categoryName, formName, fileName)) { return content; }
                using (StreamReader sr = new StreamReader(fileName, Encoding.Default))
                {
                    content = sr.ReadToEnd();
                }
                File.Delete(fileName);
            }
            catch (CommenceCOMException) { }
            return content;
        }

        private IEnumerable<CommenceViewMetaData> GetViewMetaData(IEnumerable<string> views)
        {
            foreach (string viewName in views)
            {
                IViewDef d = _db.GetViewDefinition(viewName);
                CommenceViewMetaData v = new CommenceViewMetaData(viewName, d);
                yield return v;
            }
        }

        private IEnumerable<CommenceFieldMetaData> GetFieldMetaData(string categoryName, IEnumerable<string> fields)
        {
            foreach (string fieldName in fields)
            {
                ICommenceFieldDefinition d = _db.GetFieldDefinition(categoryName, fieldName);
                CommenceFieldMetaData f = new CommenceFieldMetaData(fieldName, d);
                yield return f;
            }
        }

        private IList<FormXmlMetaData> PopulateXmlFormFilesList()
        {
            // parse all XML files belonging to Item Detail Forms
            string path = _db.Path + TEMPLATES_FOLDER;
            string[] files = Directory.GetFiles(path, "fm*.xml");
            IList<FormXmlMetaData> retval = new List<FormXmlMetaData>();
            foreach (string f in files)
            {
                using (FileStream fs = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (XmlReader reader = XmlReader.Create(fs)) //XmlReader is a fast, read-only, forward-only XML reader
                    {
                        while (reader.Read())
                        {
                            if (reader.Name.ToLower() == "form") //see if we're at the FORM node
                            {
                                FormXmlMetaData item = new FormXmlMetaData
                                {
                                    Path = f
                                };
                                if (reader.AttributeCount > 0) //only interested in the start tag; we could achieve the same with IsStartElement but that takes longer
                                {
                                    reader.MoveToAttribute("Name");
                                    //_key = reader.ReadContentAsString();
                                    item.Name = reader.ReadContentAsString();
                                    reader.MoveToAttribute("CategoryName");
                                    item.Category = reader.ReadContentAsString();
                                    retval.Add(item);
                                    break;
                                }
                            } // if form
                        } // while
                    } // using xmlreader
                } // using filestream
            } // foreach
            return retval;
        }
        #endregion

        #region Properties
        private MetaDataOptions MetaDataOptions { get; set; }
        public IEnumerable<FormXmlMetaData> FormXmlFiles
        {
            get
            {
                if (_formXmlFiles == null)
                {
                    _formXmlFiles = PopulateXmlFormFilesList();
                }
                return _formXmlFiles;
            }
        }
        #endregion

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~MetaDataBuilder() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
