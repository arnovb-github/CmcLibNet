using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes members of the CursorFilters class. For code examples see <see cref="CursorFilters"/>.
    /// </summary>
    [ComVisible(true)]
    [Guid("25EF6186-CEEE-4c54-9C52-A689548AD563")]
    public interface ICursorFilters
    {
        /// <summary>
        ///  Create new filter and add it to the collection.
        /// </summary>
        /// <param name="clauseNumber">Order of the filter, should be between 1-8.</param>
        /// <param name="filterType">The type of filter to create.</param>
        /// <returns>Derived BaseFilter corresponding to either of
        /// <see cref="ICursorFilterTypeF"/>, <see cref="ICursorFilterTypeCTI"/>, 
        /// <see cref="ICursorFilterTypeCTCF"/>, <see cref="ICursorFilterTypeCTCTI"/>.</returns>
        dynamic Create(int clauseNumber, FilterType filterType); // should return only applicable filter type.
        /// <summary>
        /// Add a filter to the filter collection.
        /// </summary>
        /// <param name="filter"><see cref="ICursorFilter"/> object.</param>
        [ComVisible(false)]
        void Add(ICursorFilter filter);
        /// <summary>
        /// Enable enumeration.
        /// </summary>
        /// <returns>ICursorFilter.</returns>
        [ComVisible(false)]
        IEnumerator<ICursorFilter> GetEnumerator();
        /// <summary>
        /// Get filter from collection based on clause number.
        /// </summary>
        /// <param name="clauseNumber">ClauseNumber.</param>
        /// <returns>CursorFilter.</returns>
        object GetFilterByClauseNumber(int clauseNumber); 
        /// <summary>
        /// Clear all filters from collection.
        /// The underlying cursor will be updated, there is no need to call <see cref="Apply()"/>.
        /// </summary>
        void Clear();
        /// <summary>
        /// Remove specified filter from collection.
        /// The underlying cursor will be updated, there is no need to call <see cref="Apply()"/>.
        /// </summary>
        /// <param name="cf">Filter to remove.</param>
        /// <returns><c>true</c> on success, <c>false</c> on error.</returns>
        bool RemoveFilter(BaseCursorFilter cf);
        /// <summary>
        /// Remove filter from collection by clause number, if present.
        /// The underlying cursor is also updated, there is no need to call <see cref="Apply()"/>.
        /// </summary>
        /// <param name="clauseNumber">Clause number (1-8).</param>
        /// <returns><c>true</c> on success, <c>false</c> on error.</returns>
        bool RemoveFilterByClauseNumber(int clauseNumber);
        /// <summary>
        /// Apply filters and return number of affected items.
        /// </summary>
        /// <returns>Number of items in cursor after applying filters.</returns>
        int Apply();
        /// <summary>
        /// Gets number of filters.
        /// </summary>
        int Count { get; }
        /// <summary>
        /// Returns filter from collection.
        /// </summary>
        /// <param name="index">Index of filter (0-based).</param>
        /// <returns>Derived BaseFilter object.</returns>
        object GetFilter(int index);
        /// <summary>
        /// Summarizes filters set so far. Useful for COM Interop clients.
        /// </summary>
        /// <returns>Single string of filterstrings for all filters in the collection.</returns>
        string ToString();
        /// <summary>
        /// Checks if filter will work.
        /// </summary>
        /// <remarks>This is a very expensive call, it creates a dummy cursor and actually tries if the filter works.</remarks>
        /// <param name="filter">ICursorFilter.</param>
        /// <returns><c>True</c> if filter succeeded, otherwise <c>false</c>.</returns>
        [ComVisible(false)]
        bool ValidateFilter(ICursorFilter filter); // cannot be bothered making this COM-visible
    }
}
