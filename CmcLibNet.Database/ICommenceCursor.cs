using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// CommenceCursor interface.
    /// </summary>
    [ComVisible(true)]
    [Guid("4C35C430-3EB7-4c6d-851A-315B0FAEFBF5")]
    public interface ICommenceCursor : ICursorEvents, IDisposable
    {
        /// <summary>
        /// Name of the category this cursor is based on.
        /// </summary>
        string Category { get; }

        /// <summary>
        /// Name of the View if cursor is based on a view.
        /// </summary>
        string View { get; }

        /// <summary>
        /// Allows for setting multiple direct and related columns at once.
        /// </summary>
        /// <remarks>Do not use this in conjunction with <c>SetColumn</c> or <c>SetRelatedColumn</c> methods. Set all columns at once using this property, or set them individually.
        /// <para>Only available for .NET applications.</para></remarks> 
        [ComVisible(false)]
        ICursorColumns Columns { get; }

        /// <summary>
        /// Number of items in cursor.
        /// </summary>
        int RowCount { get; }

        /// <summary>
        /// Number of columns in cursor.
        /// </summary>
        int ColumnCount { get; }

        /// <summary>
        /// Maximum field size. Gets the maximum number of characters that can be read from a Commence field. Undocumented by Commence.
        /// The default value is 93750.
        /// </summary>
        /// <remarks>Be careful with setting this to a large value. It obviously impacts memory usage, and may make Commence go unresponsive for a long time,
        /// especially when you have columns defined with <see cref="SetRelatedColumn"/>!
        /// <para>This property is not documented by Commence.</para></remarks>
        int MaxFieldSize { get; set; }

        /// <summary>
        /// Maximum number of rows a cursor can contain. Undocumented by Commence. Effectively a measure for database capacity.
        /// </summary>
        /// While settable in the Commence API, setting it has no effect, so in this assembly it is read-only.
        /// <remarks>At least up until Commence version RM 6.0, a Commence category can hold up to 500.000 items.</remarks>
        int MaxRows { get; }

        /// <summary>
        /// Seek to row in cursor.
        /// </summary>
        /// <param name="bkOrigin">Starting row.</param>
        /// <param name="nRows">Number of rows to move pointer.</param>
        /// <returns>Number of rows moved, -1 on error.</returns>
        int SeekRow(CmcCursorBookmark bkOrigin, int nRows);

        /// <summary>
        /// Move a fractional number of rows.
        /// </summary>
        /// <param name="nNumerator">Numerator.</param>
        /// <param name="nDenominator">Denominator.</param>
        /// <returns>Number of rows moved, -1 on error.</returns>
        int SeekRowApprox(int nNumerator, int nDenominator);

        /// <summary>
        /// Set active date used for view cursors using a view linking filter.
        /// </summary>
        /// <param name="sDate">Date (string).</param>
        /// <param name="flags">Option flags, must be 0.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        bool SetActiveDate(string sDate, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Set active date range used for view cursors using a view linking filter.
        /// </summary>
        /// <param name="startDate">Date (string)</param>
        /// /// <param name="endDate">Date (string)</param>
        /// <param name="flags">Option flags, must be 0.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        bool SetActiveDateRange(string startDate, string endDate, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Set active item used for view cursors using a view linking filter.
        /// </summary>
        /// <param name="pCategoryName">Category name of the active item used with view linking filter.</param>
        /// <param name="pRowID">Unique ID string obtained from GetRowID() indicating the active item used with view linking filter.</param>
        /// <param name="flags">Unused, must be 0.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        bool SetActiveItem(string pCategoryName, string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Defines the column set for the cursor
        /// </summary>
        /// <param name="nColumn">The (0-based) index of the column to set.</param>
        /// <param name="pName">Name of the field to use in this column.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// <see cref="CmcOptionFlags.All"/> - create column set of all fields.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        /// <remarks>Do not skip or duplicate indexes. Related columns must come after the direct columns.
        /// <para>Using <see cref="SetColumns"/> offers a more convenient way to set columns.</para></remarks>
        bool SetColumn(int nColumn, string pName, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Allows for setting direct columns in batch-mode
        /// </summary>
        /// <param name="columnNames">Object array of columnnames.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        /// <remarks>Note the <c>object</c> and not <c>string</c> type of the array, this is to accomodate COM clients.</remarks>
        bool SetColumns(object[] columnNames);

        //bool SetRelatedColumns(List<IRelatedColumn> relatedColumns);

        /// <summary>
        /// Gets collection of filters to be applied on cursor.
        /// It is recommended you use this property over setting filterstrings directly.
        /// </summary>
        ICursorFilters Filters { get; } // recommended way of filtering!!

        /// <summary>
        /// Defines the filter logic for the cursor.
        /// Not recommended, use <see cref="CommenceCursor.Filters"/> instead. 
        /// </summary>
        /// <param name="pLogic">Fully formatted [ViewConjunction()] DDE request command.</param>
        /// <param name="flags">Unused, must be <c>CmcOptionFlags.Default</c></param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        bool SetLogic(string pLogic, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Defines the filter for the cursor.
        /// Not recommended, use <see cref="CommenceCursor.Filters"/> instead. 
        /// </summary>
        /// <param name="pFilter">Fully formatted [ViewFilter()] DDE request command.</param>
        /// <param name="flags">Unused, must be <see cref="CmcOptionFlags.Default"/></param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        /// <seealso cref="Filters"/>
        bool SetFilter(string pFilter, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Define column for a connected field. Allows for retrieving other than just the Name field from connections.
        /// </summary>
        /// <param name="nColumn">The (0-based) index of the column to set.</param>
        /// <param name="pConnName">Name of the connection to use in this column.</param>
        /// <param name="pCatName">Name of the connected Category to use in this column.</param>
        /// <param name="pName">Name of the field in the connected category to use in this column</param>
        /// <param name="flags">Logical OR of following option flags:
        /// <see cref="CmcOptionFlags.All"/> - create column set of all fields.</param>
        /// <remarks>Setting the CmcOptionFlags.All flag *only* affects direct fields in the cursor.
        /// It cannot be used to obtain all fields from the connected category.</remarks>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        /// <remarks>Related columns must always come after the direct columns. Do not skip or duplicate an index.
        /// <para>Warning: related columns severely impact memory usage and processing time!</para>
        /// <para>Using <c>Columns</c> offers a more convenient way to set related columns.</para></remarks>
        bool SetRelatedColumn(int nColumn, string pConnName, string pCatName, string pName, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Defines the sort criteria for the cursor. Use this only if you apply filters with the SetFilter method. You are strongly advised to use the Filters collection instead.
        /// </summary>
        /// <param name="pSort">Text defining the new sort criteria. Syntax is identical to the one used by the DDE ViewSort request.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns>Returns <c>true</c> on succes, <c>false</c> on error.</returns>
        bool SetSort(string pSort, CmcOptionFlags flags);

        /// <summary>
        /// (read-only) TRUE if category is shared in a workgroup.
        /// </summary>
        bool Shared { get; }

        /// <summary>
        /// Create a rowset of new items to add to the database.
        /// </summary>
        /// <param name="nRows">Number of rows to create.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// <see cref="CmcOptionFlags.Shared"/> - all rows default to shared.</param>
        /// <returns><see cref="ICommenceAddRowSet"/> object.</returns>
        CmcLibNet.Database.ICommenceAddRowSet GetAddRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset of existing items for editing.
        /// </summary>
        /// <param name="nRows">Number of rows to retrieve.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceEditRowSet"/> object.</returns>
        CmcLibNet.Database.ICommenceEditRowSet GetEditRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset of items in cursor for editing. Be careful when handling very large rowsets.
        /// </summary>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceEditRowSet"/> object on success.</returns>
        /// <remarks>This method is only available to .NET applications.</remarks>
        [ComVisible(false)] // overload, only available for .Net
        CmcLibNet.Database.ICommenceEditRowSet GetEditRowSet(CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset for editing a specified cursor row (i.e. item).
        /// </summary>
        /// <param name="pRowID">Unique ID string obtained from GetRowID().</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceEditRowSet"/> object.</returns>
        CmcLibNet.Database.ICommenceEditRowSet GetEditRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset object with the results of a query.
        /// </summary>
        /// <param name="nRows">Maximum number of rows to retrieve.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceQueryRowSet"/> object on success.</returns>
        CmcLibNet.Database.ICommenceQueryRowSet GetQueryRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset object with the results of a query.
        /// </summary>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <remarks>This method is only available to .NET applications.</remarks>
        /// <returns><see cref="ICommenceQueryRowSet"/> object on success.</returns>
        [ComVisible(false)] // overload, only available for .Net
        CmcLibNet.Database.ICommenceQueryRowSet GetQueryRowSet(CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset object with a particular row loaded.
        /// </summary>
        /// <param name="pRowID">Unique ID string obtained from <see cref="ICommenceQueryRowSet.GetRowID"/>.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceQueryRowSet"/> on success.</returns>
        CmcLibNet.Database.ICommenceQueryRowSet GetQueryRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset object with a particular thid loaded. Only applies to shared items. Cursor must have the thids <see cref="CmcOptionFlags.UseThids">thids</see> flag.
        /// </summary>
        /// <param name="pThid">THID parts "thidId : thidSequence" (i.e., omit the leading category part).</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceQueryRowSet"/> on success</returns>
        /// <remarks>Undocumented by Commence.</remarks>
        CmcLibNet.Database.ICommenceQueryRowSet GetQueryRowSetByThid(string pThid, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset of n items for deletion.
        /// </summary>
        /// <param name="nRows">Number of rows to retrieve.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceDeleteRowSet"/> on success.</returns>
        CmcLibNet.Database.ICommenceDeleteRowSet GetDeleteRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset of ALL existing items for deletion. Use with care!
        /// </summary>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceDeleteRowSet"/> on success.</returns>
        /// <remarks>This method is only available to .NET applications.</remarks>
        [ComVisible(false)] // overload, only available for .Net
        CmcLibNet.Database.ICommenceDeleteRowSet GetDeleteRowSet(CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Create a rowset for deleting a particular row.
        /// </summary>
        /// <param name="pRowID">Unique ID string obtained from GetRowID().</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns><see cref="ICommenceDeleteRowSet"/> on success.</returns>
        CmcLibNet.Database.ICommenceDeleteRowSet GetDeleteRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default);

        /// <summary>
        /// Checks if field contains duplicate values.
        /// </summary>
        ///<param name="columnName">Columname in cursor to compare. Column must exist in the cursor.</param>
        /// <param name="caseSensitive">Perform case-sensitive check, default is <c>true</c>.</param>
        /// <returns><c>true</c> if duplicates are found.</returns>
        /// <exception cref="CommenceCOMException">Column not present in cursor.</exception>
        bool HasDuplicates(string columnName, bool caseSensitive = true);

        /// <summary>
        /// Exports current cursor to a file. WARNING: The cursor is disposed after this method!
        /// </summary>
        /// <param name="fileName">(Fully qualified) filename. Overwrites existing file.</param>
        /// <param name="settings"><see cref="Export.IExportSettings"/> object.</param>
        /// <remarks>
        /// The <see cref="Export.IExportSettings.PreserveAllConnections"/> options is ignored in this method.
        /// Alternatively, but extremely slow, you can use either the <see cref="Export.IExportSettings.UseDDE"/> option or crank up the 
        /// <see cref="Export.IExportSettings.MaxFieldSize"/> value and hope you have enough RAM memory.</remarks>
        void ExportToFile(string fileName, Export.IExportSettings settings = null);

        /// <summary>
        /// Reads specified row from the cursor.
        /// <para>This method is just a convenient way to return data without having to create a <see cref="ICommenceQueryRowSet"/>.</para>
        /// <para>If you need to apply formatting, see the <see cref="Export.IExportEngine"/> interface.</para>
        /// </summary>
        /// <remarks>This method is only available in .NET.
        /// <para>If the cursor has the <see cref="CmcOptionFlags.UseThids"/> flag defined, the first element of the inner list will contain the thid.
        /// In that case the inner list count will be the number of columns in the category plus one.</para>
        /// </remarks>
        /// <param name="lRow">Row to read. Row is 0-based, from beginning of cursor.</param>
        /// <returns>List of rowvalues, <c>null</c> on error.</returns>
        [ComVisible(false)]
        List<string> ReadRow(int lRow);

        /// <summary>
        /// Read all data from a cursor.
        /// <para>This method is just a convenient way to return data without having to create a <see cref="ICommenceQueryRowSet"/>.</para>
        /// <para>Keep in mind reading data from Commence is quite slow!</para>
        /// <para>For more advanced reading and export options see <see cref="Export.IExportEngine"/>.</para>
        /// </summary>
        /// <returns>List of lists containing commence rowvalues, <c>null</c> on error.</returns>
        /// <remarks>This method is not available to COM clients.
        /// <para>If the cursor has the <see cref="CmcOptionFlags.UseThids"/> flag defined, the first element of the inner list will contain the thid.
        /// In that case the inner list count will be the number of columns in the category plus one.</para>
        /// </remarks>
        /// <param name="batchRows">Number of rows to read per iteration.</param>
        [ComVisible(false)]
        List<List<string>> ReadAllRows(int batchRows = 1000);

        /// <summary>
        /// Get columnames from cursor. Defaults to fieldnames
        /// </summary>
        /// <param name="flags"><see cref="CmcOptionFlags"/>, defaults to <see cref="CmcOptionFlags.Fieldname"/></param>
        /// <returns>List of columnames.</returns>
        /// <remarks><para>This method is only available in .NET.</para> 
        /// <para>Pass <see cref="CmcOptionFlags.Default"/> to retrieve fieldlabels instead of fieldnames.</para></remarks>
        [ComVisible(false)]
        IEnumerable<string> GetColumnNames(CmcOptionFlags flags = CmcOptionFlags.Fieldname);

        /// <summary>
        /// Close any references to Commence. The object should be disposed after this.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references and/or release memory. It can be useful in some cases, because Commence may complain about running out of memory before the Garbage Collector has a chance to kick in.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
    }
}