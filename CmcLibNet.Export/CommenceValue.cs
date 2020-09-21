namespace Vovin.CmcLibNet.Export
{


    /// <summary>
    /// Represents a single Commence rowvalue.
    /// </summary>
    internal class CommenceValue
    {
        /* This class holds the raw commence data for a requested cursor row and its associated columndefinition.
        * Its purpose is to provide a consistent way of representing rowdata throughout the assembly.
        *
        * It is up to the consumer to get either the DirectFieldValue (=direct fieldvalue)
        * or the ConnectedValues (=fieldvalues from connections).
        * This information is contained in the ColumnDefinition.
        * 
        * This approach is overly complex, we could have just a single property, being array with values.
        * if direct, it would contain 1 element, if connected 0, 1 or many.
        * We could also get rid of the IsEmpty property.
        * 
        * It does not matter that much, because from the value itself it is not possible to tell what kind of field we are dealing with.
        * We need to query ColumnDefinition anyway, so there isn't much gained when using just a single value property.
        */

        #region Constructors
        /// <summary>
        /// Used for single values.
        /// </summary>
        /// <param name="singlevalue">single fieldvalue.</param>
        /// <param name="cd"><see cref="ColumnDefinition"/>.</param>
        internal CommenceValue(string singlevalue, ColumnDefinition cd)
        {
            DirectFieldValue = singlevalue;
            this.ColumnDefinition = cd;
        }
        /// <summary>
        /// Used for connected fields.
        /// </summary>
        /// <param name="connectedvalues">Array of connected fields.</param>
        /// <param name="cd"><see cref="ColumnDefinition"/>.</param>
        internal CommenceValue(string[] connectedvalues, ColumnDefinition cd)
        {
            this.ColumnDefinition = cd;
            ConnectedFieldValues = connectedvalues;
        }

        /// <summary>
        /// Used when there are no field values but we still want an instance.
        /// </summary>
        /// <param name="cd"><see cref="ColumnDefinition"/>.</param>
        internal CommenceValue(ColumnDefinition cd)
        {
            this.ColumnDefinition = cd;
        }
        #endregion

        #region Properties

        /// <summary>
        /// Direct field.
        /// </summary>
        internal string DirectFieldValue // direct value
        { get; } = null;

        /// <summary>
        /// String array containing connected values.
        /// </summary>
        internal string[] ConnectedFieldValues { get; } = null;

        /// <summary>
        /// Columndefinition of the field.
        /// </summary>
        internal ColumnDefinition ColumnDefinition { get; private set; }

        /// <summary>
        /// Returns true if there is no fieldvalue.
        /// </summary>
        internal bool IsEmpty
        {
            get
            {
                return (this.DirectFieldValue == null && this.ConnectedFieldValues == null);
            }
        }

        #endregion

    }
} 
