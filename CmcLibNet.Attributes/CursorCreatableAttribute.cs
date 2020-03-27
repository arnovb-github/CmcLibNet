using System;

namespace Vovin.CmcLibNet.Attributes
{
    /// <summary>
    /// Indicates if a Commence cursor can be created on the Commence View type.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field)]
    public class CursorCreatableAttribute : Attribute
    {
        #region Properties

        /// <summary>
        /// Property.
        /// </summary>
        public bool CursorCreatable { get; protected set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="value">Boolean.</param>
        public CursorCreatableAttribute(bool value)
        {
            this.CursorCreatable = value;
        }

        #endregion
    }
}
