using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Contains the category definition.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("ABA6E0E8-F97B-4eb5-877A-E2C9C86BEF21")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICategoryDef))]
    public class CategoryDef : ICategoryDef
    {
        private int _categoryID = -1;
        internal CategoryDef() { }
        /// <inheritdoc />
        public int MaxItems { get; internal set; }
        /// <inheritdoc />
        public bool Shared { get; internal set; }
        /// <inheritdoc />
        public bool Duplicates { get; internal set; }
        /// <inheritdoc />
        public bool Clarified { get; internal set; }
        /// <inheritdoc />
        public string ClarifySeparator { get; internal set; }
        /// <inheritdoc />
        public string ClarifyField { get; internal set; }
        /// <inheritdoc />
        public int CategoryID { 
            get
            {
                return _categoryID;
            }
            internal set
            {
                _categoryID = value;
            }
        }
    }
}
