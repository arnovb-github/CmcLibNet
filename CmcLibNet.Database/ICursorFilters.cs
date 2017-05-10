using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes members of the CursorFilters class. For code examples see <see cref="CursorFilters"/>.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("25EF6186-CEEE-4c54-9C52-A689548AD563")]
    public interface ICursorFilters
    {
        /// <summary>
        ///  Create new filter and add it to the collection.
        /// </summary>
        /// <param name="clauseNumber">clauseNumber is the order of the filter, should be between 1-8.</param>
        /// <param name="filterType">The type of filter to create.</param>
        /// <returns>Derived BaseFilter corresponding with filtertype.</returns>
        dynamic Create(int clauseNumber, FilterType filterType); // should return only applicable filter type
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
        /// <param name="cf">CursorFilter.</param>
        /// <returns><c>true</c> on success, <c>false</c> on error.</returns>
        bool RemoveFilter(CursorFilter cf);
        /// <summary>
        /// Remove filter from collection by clause number, if present.
        /// The underlying cursor is also updated, there is no need to call <see cref="Apply()"/>.
        /// </summary>
        /// <param name="clauseNumber">Clause number (1-8).</param>
        /// <returns><c>true</c> on success, <c>false</c> on error.</returns>
        bool RemoveFilterByClauseNumber(int clauseNumber);
        /// <summary>
        /// Apply the specified filters.
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
    }
}
