using Newtonsoft.Json;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Database
{
    // this portion of the class deals with all the Schema information
    public partial class CommenceDatabase : ICommenceDatabase
    {
        /// <inheritdoc />
        public IDatabaseSchema GetDatabaseSchema(MetaDataOptions options = null)
        {
            if (options == null) { options = new MetaDataOptions(); }
            using (MetaDataBuilder mb = new MetaDataBuilder(this, options))
            {
                return mb.BuildDatabaseSchema();
            }
        }

        /// <inheritdoc />
        public void ExportDatabaseSchema(string fileName, MetaDataOptions options = null)
        {
            if (options == null)
            {
                options = new MetaDataOptions();
            }
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("Filename null or empty.", nameof(fileName));
            }

            using (ICommenceDatabase db = new CommenceDatabase())
            {
                var schema = db.GetDatabaseSchema(options);

                switch (options.Format)
                {
                    default:
                    case MetaDataFormat.Json:
                        var s = JsonConvert.SerializeObject(schema);
                        using (StreamWriter sw = new StreamWriter(fileName))
                        {
                            sw.Write(s);
                        }
                        break;
                    case MetaDataFormat.Xml:

                        XmlSerializer xsSubmit = new XmlSerializer(typeof(DatabaseSchema));
                        var xml = string.Empty;

                        using (var sw = new StreamWriter(fileName))
                        {
                            using (XmlWriter writer = XmlWriter.Create(sw))
                            {
                                xsSubmit.Serialize(writer, schema);
                                xml = sw.ToString(); // Your XML
                            }
                        }
                        break;
                }
            }
        }
    }
}
