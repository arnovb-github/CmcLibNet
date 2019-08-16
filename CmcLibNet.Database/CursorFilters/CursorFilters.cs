using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Class for creating, removing and applying Cursor filters.
    /// <para>This is significantly different from the usual way of filtering in Commence API.</para>
    /// <para>See code samples below.</para>
    /// </summary>
    /// <example>
    /// <para>Filtering cursors is a very complex affair in he Commence API, because it is done using DDE-style syntax. This syntax is convoluted and hard to remember.</para>
    /// <para>Old-style DDE syntax can still be used in CmcLibNet, but you are strongly encouraged to use the CursorFilters class, exposed as Filters property of the <c>CommenceCursor</c> class. It is more verbose, but also more intuitive.</para>
    /// <para>First you decide what type of filter you want to apply. Commence has 4 types of filters, which are defined in the <see cref="FilterType"/> enum.
    /// Once you know what filter you want to apply, you create it:</para>
    /// <code>
    /// using Vovin.CmcLibNet
    /// using Vovin.CmcLibNet.Database
    /// 
    /// namespace FilterExample
    /// {
    /// 	class Main
    /// 		{
    /// 		    // get a reference to CmcNetLib.
    /// 			ICommenceDatabase cmc = new CommenceDatabase(); 
    /// 			// get a cursor from Commence.
    /// 			ICommenceCursor cur = cmc.GetCursor("Account");
    /// 			// get a Field type filter
    /// 			ICursorFilterTypeF f = cur.Filters.Add(1, FilterType.Field);
    /// 		}
    /// }
    /// </code>
    /// 
    /// <para>Step-through of above example: First, we get a reference to the CmcLibNet assembly, that will talk to Commence for you.</para>
    /// <para>Then, we get a cursor, telling it we want to use data from the Account category. We want to filter that category on a fieldvalue.</para>
    /// <para>We then define our filter as ICursorFilterTypeF.</para>
    /// <para>We then ask CmcLibNet to get use a filter of that type, the 'FilterType.Field' parameter, and make it the first filter (the '1' parameter).
    /// The index number is mandatory, because of the way the and/or relations between filters in Commence work.</para>
    /// 
    /// <para>What happens is a filter object instance called 'f' of type CursorFilterTypeF is created.</para> 
    /// <para>It just exposes the properties pertaining to that particular filter-type. Have peek with Intellisense or the object-browser and you'll see what I mean. So now, you can set the filter parameters. Most are optional. Let's set them all (well, almost all):</para>
    /// 
    /// <code>
    /// f.FieldName = "accountKey"; // field to filter on
    /// f.FieldValue = "Commence"; // value to filter on
    /// f.Qualifier = FilterQualifier.Contains; // the Qualifier property sets the proper QualifierString internally.
    /// f.CSFlag = true; // tell Commence we want to do a case-sensitive filter. You can omit this for a non-case-sensitive filter.
    /// f.Except = true; // get us all results EXCEPT the ones matching our filter. This is equivalent to the checking the 'except' checkbox in the filter dialog in Commence
    /// f.OrFlag = false; // we do not want to define this filter as an OR filter. Note that the filter logic is incorporated as a property of the individual filters, and not supplied separately by a SetLogic method.
    /// </code>
    /// 
    /// This equates to the shorter:
    ///  <code>
    ///  cur.SetFilter("[ViewFilter(1,NOT,F,Contains," + "Commence" + ",1)]",0);
    ///  cur.SetLogic("[ViewConjunction(And,,,,,)]",0);
    ///  </code>but is is easier to read and understand as well as less error-prone.
    /// 
    /// <para>Note that the filter logic is incorporated as a property of the individual filters, and not supplied separately by a SetLogic method.</para>
    /// 
    /// <para>This was just 1 filter, you can set up to 8 of them. Once you are done with defining your filters, call the Apply method of the Filters class:</para>
    /// 
    /// <code>cur.Filters.Apply();</code>
    /// 
    /// That's it for C# users.
    /// </example>
    /// <example>
    /// For VBScript (e.g. an Item Detail Form Script), the procedure is much alike.
    /// The main difference is that it is probaly easiest to supply the filter qualifier as a string via QualifierString.
    /// (Although if you want to, you can supply the enum value to Qualifier):
    /// 
    /// <code language="vbscript">
    /// Dim cmc : Set db = CreateObject("CmcNetLib.Database")
    /// Dim cur : Set cur = db.GetCursor("Account")
    /// Dim filters : Set filters = cur.Filters ' you cannot simply use cur.Filters from VBScript, alas.
    /// Dim f : Set f = filters.Add(1, 0) ' 0 is the enum value for FilterType.Field.
    /// f.FieldName = "accountKey" ' field to filter on.
    /// f.FieldValue = "Commence" ' value to filter on.
    /// f.QualifierString = "contains" ' the qualifier, i.e., how to evaluate the filter value. You could have also used f.Qualifier with the enum value corresponding with 'contains', this is just a bit easier.
    /// f.CSFlag = true ' tell Commence we want to do a case-sensitive filter. You can omit this for a non-case-sensitive filter.
    /// f.Except = true ' get us all results EXCEPT the ones matching our filter. This is equivalent to the checking the 'except' checkbox in the filter dialog in Commence.
    /// f.OrFlag = false ' we do not want to define this filter as an OR filter. 
    /// filters.Apply()
    /// '...do something...
    /// db.Close
    /// </code>
    /// </example>

    [ComVisible(true)]
    [Guid("57C88F8C-8D4C-4ba3-9487-4354065809D4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilters))]
    public class CursorFilters : ICursorFilters, IEnumerable<ICursorFilter>
    {
        private Database.ICommenceCursor _cur = null;
        private List<ICursorFilter> _filters = new List<ICursorFilter>();
        private const int _MAX_FILTERS = 8;

        #region Constructors
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="cursor">Database.ICommenceCursor object.</param>
        internal CursorFilters(Database.ICommenceCursor cursor)
        {
            _cur = cursor;
        }
        #endregion

        #region Methods

        /// <inheritdoc />
        // TODO rethink this
        public dynamic Add(int clauseNumber, FilterType filterType) // should we overload this somehow? What's the best way? Remember that methods with the same parameter signature cannot be overloaded!
        {
            // Note the dynamic return type. COM Interop requires an object;
            // it cannot be the base class, because in that case COM Interop only exposes the base interface.
            // and then .NET users have to use an explicit cast to get the type right, which is counterintuitive.
            // There is no way to solve this using MarshalAs as far as I know
            // However, we can sort of get away with using the 'dynamic' keyword.
            // .Net consumers will still need to cast the type, but at least it should be 'visible'(?verify!)

            if (_filters.Count() == _MAX_FILTERS)
            {
                throw new IndexOutOfRangeException("Maximum number of filters in use (" + _filters.Count().ToString() + ").");
            }

            switch (filterType)
            {
                case FilterType.Field:
                    ICursorFilterTypeF f = new CursorFilterTypeF(clauseNumber);
                    _filters.Add(f);
                    return f;
                case FilterType.ConnectionToCategoryField:
                    ICursorFilterTypeCTCF ctcf = new CursorFilterTypeCTCF(clauseNumber);
                    _filters.Add(ctcf);
                    return ctcf;
                case FilterType.ConnectionToItem:
                    ICursorFilterTypeCTI cti = new CursorFilterTypeCTI(clauseNumber);
                    _filters.Add(cti);
                    return cti;
                case FilterType.ConnectionToCategoryToItem:
                    ICursorFilterTypeCTCTI ctcti = new CursorFilterTypeCTCTI(clauseNumber);
                    _filters.Add(ctcti);
                    return ctcti;
                default:
                    return null; // not pretty
            }
        }

        /// <inheritdoc />
        public object GetFilterByClauseNumber(int clauseNumber)
        {
            return _filters.SingleOrDefault(w => w.ClauseNumber == clauseNumber);
        }

        /// <inheritdoc />
        public object GetFilter(int index)
        {
            if ((index > -1) && (index < _filters.Count))
            {
                return _filters.ElementAt(index);
            }
            else
            {
                return null;
            }
        }

        /// <inheritdoc />
        public void Clear()
        {
            foreach(BaseCursorFilter f in _filters)
            {
                string s = "[ViewFilter(" + f.ClauseNumber.ToString() + ",\"Clear\")]";
                if (_cur.SetFilter(s, CmcOptionFlags.Default) == false)
                {
                    // throwing an error exits a foreach loop
                    throw new CommenceCOMException("Failed to clear filter '" + s + "' on cursor.\n\nYou should recreate the cursor.");
                }
            }
            _filters.Clear();
        }

        /// <inheritdoc />
        public bool RemoveFilter(BaseCursorFilter cf)
        {
            string s = "[ViewFilter(" + cf.ClauseNumber.ToString() + ",\"Clear\")]";
            if (_cur.SetFilter(s, CmcOptionFlags.Default) == false)
            {
                throw new CommenceCOMException("Cursor method SetFilter failed on filter: '" + s + "'.\nYou should recreate the cursor.");
            }
            _filters.Remove(cf);
            return true;
        }

        /// <inheritdoc />
        public bool RemoveFilterByClauseNumber(int clauseNumber)
        {
            var f = _filters.SingleOrDefault(w => w.ClauseNumber == clauseNumber);
            string s = "[ViewFilter(" + f.ClauseNumber.ToString() + ",\"Clear\")]";
            if (_cur.SetFilter(s, CmcOptionFlags.Default) == false)
            {   
                throw new CommenceCOMException("Cursor method SetFilter failed on filter: '" + s + "'.\nYou should recreate the cursor.");
            }
            return _filters.Remove(f);
        }

        /// <inheritdoc />
        public int Apply()
        {
            // evaluate and apply filters
            StringBuilder sb = null;
            List<string> logic = new List<string>();

            // sort the list by clausenumber to make sure the order of the logic is set correctly.
            List<ICursorFilter> sortedList = _filters.OrderBy(o => o.ClauseNumber).ToList();

            foreach (ICursorFilter f in sortedList)
            {
                if (sb == null) { sb = new StringBuilder("[ViewConjunction("); }
                if (_cur.SetFilter(f.ToString(), CmcOptionFlags.Default) == false)
                {
                    /* Applying the filter failed.
                     * We can do several things now. We could catch the error and return an error value like false or -1.
                     * But then the user wouldn't know which filter failed.
                     * How about the filters that were already applied?
                     * Should we clear those just in case?
                     * Or would simply calling Apply again (after fixing the filter) be enough?
                     * i.o.w should we make CommenceCursor go out of scope or not?
                     * it would make things horribly complicated to understand?
                     */
                    sb = null;
                    /* Not sure if this is the best place to throw, it might be better to throw it earlier, in CommenceCursor
                     * however, if we do it in CommenceCursor, the default behaviour would change from an Interop consumer point of view
                     * I.o.w., it would no longer comply with what Commence documentation specifies and that is unwanted
                     * throwing it here on the other hand makes this assembly harder to debug.
                     * The actual error occurs in CommenceCursor.SetFilter
                     */
                    throw new CommenceCOMException("CommenceCursor method SetFilter failed on filter request:\n\n'" + f.ToString() + "'");
                } // if

                // add logic string.
                logic.Add((f.OrFilter) ? "Or" : "And");
            } // foreach

            // successfully applied filters, now set logic
            // not very elegant, could do with refactoring? TODO
            if (sb != null)
            {
                sb.Append(string.Join(",",logic));
                sb.Append(")]");
                if (_cur.SetLogic(sb.ToString(), CmcOptionFlags.Default) == false)
                {
                    throw new CommenceCOMException("CommenceCursor method SetLogic failed on logic request:\n\n'" + sb.ToString() + "'");
                }
            }
            return _cur.RowCount; // return number of filtered rows
        }

        /// <inheritdoc />
        public override string ToString()
        {
            // displays all filter clauses
            StringBuilder sb = new StringBuilder();
            foreach (ICursorFilter f in _filters)
            {
                sb.AppendLine(f.GetViewFilterString());
            }
            return sb.ToString();
        }
        /// <summary>
        /// Method to check if filterclause is in use.
        /// </summary>
        /// <param name="filters">List of filters.</param>
        /// <param name="clauseNumber">Clause number to check.</param>
        /// <returns><c>true</c> if clausenumber in use, otherwise <c>false</c>.</returns>
        private bool ClauseInUse(List<ICursorFilter> filters, int clauseNumber)
        {
            return filters.Any(a => a.ClauseNumber == clauseNumber);
        }

        /// <summary>
        /// Get Enumerator.
        /// </summary>
        /// <returns>IEnumerator.</returns>
        public IEnumerator<ICursorFilter> GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        /// <summary>
        /// Get enumerator.
        /// </summary>
        /// <returns>System.Collections.IEnumerator.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _filters.GetEnumerator();
        }
        #endregion

        #region Properties
        /// <inheritdoc />
        public int Count
        {
            get { return _filters.Count; }
        }
        #endregion
    }
}
