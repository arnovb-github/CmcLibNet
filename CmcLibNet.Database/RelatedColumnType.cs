namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Keep track of the underlying fieldtype. Needed because different fieldtypes are returned with a different separator by Commence.
    /// </summary>
    internal enum RelatedColumnType
    {
        Connection = 0,
        ConnectedField = 1
    }
}
