using System;
using System.Reflection;

namespace Vovin.CmcLibNet.Extensions
{
    /// <summary>
    /// Enum extension methods
    /// </summary>
    internal static class EnumExtensions
    {
        /// <summary>
        /// Get value of attribute property for enum value.
        /// </summary>
        /// <typeparam name="T">Return type.</typeparam>
        /// <typeparam name="TAttr">Attribute.</typeparam>
        /// <param name="value">Enum value.</param>
        /// <param name="attrPropertyName"></param>
        /// <returns>Value of attribute property for enum value.</returns>
        internal static T GetAttributePropertyValue<T, TAttr>(this Enum value, string attrPropertyName)
            where TAttr : Attribute
        {
            Type attrType = typeof(TAttr);
            FieldInfo fi = value.GetType().GetField(value.ToString());
            PropertyInfo propInfoAttributePropertyName = attrType.GetProperty(attrPropertyName);

            if (propInfoAttributePropertyName is null)
            {
                throw new ArgumentException($"Attribute with property {attrPropertyName} not found in enum {value.ToString()}.");
            }

            TAttr[] attributes =
                (TAttr[])fi.GetCustomAttributes(attrType, false);

            if (attributes?.Length > 0)
            {
                return (T)propInfoAttributePropertyName.GetValue(attributes[0]);
            }
            else
                return default(T);
        }
    }
}
