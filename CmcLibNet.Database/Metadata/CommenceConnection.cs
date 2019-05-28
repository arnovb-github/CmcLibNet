using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Describes a Commence connection.
    /// </summary>
    [Serializable]
    [ComVisible(true)]
    [Guid("9CF29D9A-D8DF-4092-866A-01D4C1B11200")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceConnection))]
    public class CommenceConnection : ICommenceConnection
    {
        /// <inheritdoc />
        public string Name { get; set; }

        /// <inheritdoc />
        public string ToCategory { get; set; }

        /// <inheritdoc />
        public string FullName
        {
            get
            {
                if (string.IsNullOrEmpty(Name) || string.IsNullOrEmpty(ToCategory))
                {
                    throw new NullReferenceException("Name or Category not specified.");
                }
                else
                    return Name + ' ' + ToCategory;
            }
            set { }
        }
    }
}
