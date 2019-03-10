namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Abstract base class for the various *RowSet Types.
    /// </summary>
    public abstract class BaseRowSet : IBaseRowSet
    {
        #region Fields
        /// <summary>
        /// Delimiters used in Commence DDE conversations
        /// By defining them here, consumers do not have to supply them with every DDE call that uses them.
        /// </summary>
        protected internal static readonly string CMC_DELIM = @"@@##~~&&";
        /// <summary>
        /// Secondary delimiter.
        /// </summary>
        protected internal static readonly string CMC_DELIM2 = @"@@/\~_&&";
        /// <summary>
        /// String.Split requires a string array.
        /// </summary>
        protected internal readonly string[] _splitter = new string[] { CMC_DELIM };
        /// <summary>
        /// String.Split requires a string array.
        /// </summary>
        protected internal readonly string[] _splitter2 = new string[] { CMC_DELIM2 };
        // Flag: Has Dispose already been called?
        bool disposed = false;
        #endregion

        #region Constructors
        /// <summary>
        /// Internal constructor.
        /// </summary>
        protected internal BaseRowSet() { }
        /// <summary>
        /// Destructor.
        /// </summary>
        ~BaseRowSet()
        {
            Dispose(false);
        }
        #endregion

        /// <inheritdoc />
        public abstract int RowCount { get;}
        /// <inheritdoc />
        public abstract int ColumnCount { get; }
        /// <inheritdoc />
        public abstract string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        public abstract string GetColumnLabel(int nCol, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        public abstract int GetColumnIndex(string pLabel, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        public abstract object[] GetRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        public abstract bool GetShared(int nRow);
        /// <inheritdoc />
        public abstract void Close(); // just an alias for Dispose for COM users
        internal abstract void RCWReleaseHandler(object sender, System.EventArgs e);

        /// <summary>
        /// Returns the primary delimiter used in this assembly.
        /// </summary>
        protected string Delim { get { return CMC_DELIM; } }

        /// <summary>
        /// Creates object array from string array.
        /// </summary>
        /// <param name="input">string array.</param>
        /// <returns>object array that can be consumed by COM clients such as VBScript.</returns>
        protected internal static object[] toObjectArray(string[] input)
        {
            object[] objArray = new object[input.Length];
            input.CopyTo(objArray, 0);
            return objArray;
        }

        /// <summary>
        /// Dispose method
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            System.GC.SuppressFinalize(this); 
        }

        // Protected implementation of Dispose pattern.
        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }
    }
}
