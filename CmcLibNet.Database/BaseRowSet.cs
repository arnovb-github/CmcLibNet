namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Abstract base class for the various *RowSet Types.
    /// </summary>
    public abstract class BaseRowSet : IBaseRowSet
    {
        #region Fields
        /// <summary>
        /// Delimiter used in GetRow().
        /// </summary>
        private static readonly string _cmcDelim = @"@@##~~&&";
        /// <summary>
        /// For use in <see cref="System.String.Split(char[], System.StringSplitOptions)"/>
        /// </summary>
        private readonly string[] _splitter = new string[] { _cmcDelim };
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
        public abstract int RowCount { get; }
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
        public abstract object[] GetRow(int nRow, string delim, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        public abstract bool GetShared(int nRow);
        /// <inheritdoc />
        public abstract void Close(); // just an alias for Dispose for COM users
        internal abstract void RCWReleaseHandler(object sender, System.EventArgs e);

        /// <summary>
        /// Returns the primary delimiter used in this assembly.
        /// </summary>
        protected string Delim => _cmcDelim;

        /// <summary>
        ///  For use in <see cref="System.String.Split(char[], System.StringSplitOptions)"/>
        /// </summary>
        // We do not use an auto-property because it would introduce overhead
        protected string[] Splitter =>_splitter;

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
        //// we *could* bring all methods and properties that all ICommenceXRowset share to the base class
        //// by using the below implementation. Would make the code more DRY
        //// It would work but is very hard to debug because of the dynamic keyword.
        //// Also, it makes code slightly slower.
        //protected internal object[] GetRow(dynamic rs, int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        //{
        //    return ((string)rs.GetRow(nRow, Delim, (int)flags))
        //        .Split(_splitter, StringSplitOptions.None)
        //        .ToArray<object>();
        //}
    }
}
