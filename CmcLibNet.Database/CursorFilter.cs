using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// CursorFilter is a an abstract base class for the various Commence filters you can create.
    /// Use CursorFilters.Add to get a derived filter object that exposes properties only applicable to that filtertype.
    /// This shields you from having to use the horrible DDE syntax the ViewFilter request takes.
    /// You can still set filters the old-fashioned way via CommenceCursor.SetFilter.
    /// </summary>
    [ComVisible(true)]
    [Guid("691B0432-73F1-4508-B44E-A8AACB26F1FB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilter))]
    /* a problem we will now run into is,
     * that interfaces of derived classes are not fully exposed to COM
     * if early-binding is used, we're fine,
     * but when late binding is used,
     * the caller does not have access to the derived interface;
     * it will only see the ICursorFilter interface.
     * This is why we use the 'new' keyword in the derived interfaces.
     * */
    public abstract class CursorFilter : ICursorFilter
    {
        /// <summary>
        /// The clause number is a 1-based int specifying the filter order. There can be up to 8 filters.
        /// </summary>
        private int _clauseNumber;

        /// <summary>
        /// Constructor checks and sets the clausenumber.
        /// </summary>
        /// <param name="clauseNumber">Clause number, should be between 1 and 8.</param>
        protected internal CursorFilter(int clauseNumber)
        {
            ClauseNumber = clauseNumber;
        }

        #region Properties
        /// <inheritdoc />
        public bool OrFilter { get; set; }
        /// <inheritdoc />
        public bool Except { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Returns ClauseNumber for use in derived classes.
        /// Should be a value between 1 and 8.
        /// </summary>
        public int ClauseNumber
        {
            get { return _clauseNumber; }
            set
            {
                if (value >= 1 && value <= 8)
                {
                    _clauseNumber = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Clause number must be between 1 and 8");
                }
            }
        }

        /// <inheritdoc />
        public virtual string GetViewFilterString()
        {
            return this.ToString();
        }
        #endregion
    }
}
