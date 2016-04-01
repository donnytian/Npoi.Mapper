using System;
using System.Reflection;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required for one column when mapping between object and file rows.
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public class ColumnInfo<TTarget>
    {
        #region Properties

        /// <summary>
        /// Value for the column header.
        /// </summary>
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public object HeaderValue { get; private set; }

        /// <summary>
        /// The column resolver.
        /// </summary>
        public ColumnResolver<TTarget> Resolver { get; set; }

        /// <summary>
        /// The mapped property information.
        /// </summary>
        public PropertyMeta PropertyMeta { get; }

        /// <summary>
        /// The last non-blank value.
        /// </summary>
        public object LastNonBlankValue { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo{TTarget}"/> class.
        /// </summary>
        /// <param name="headerValue">The header value</param>
        /// <param name="columnName">The column name.</param>
        /// <param name="pi">The mapped PropertyInfo.</param>
        public ColumnInfo(object headerValue, string columnName, PropertyInfo pi)
        {
            HeaderValue = headerValue;
            PropertyMeta = new PropertyMeta(columnName, pi);
        }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo{TTarget}"/> class.
        /// </summary>
        /// <param name="headerValue">The header value</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="pi">The mapped PropertyInfo.</param>
        public ColumnInfo(object headerValue, int columnIndex, PropertyInfo pi)
        {
            HeaderValue = headerValue;
            PropertyMeta = new PropertyMeta(columnIndex, pi);
        }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo{TTarget}"/> class.
        /// </summary>
        /// <param name="headerValue">The header value</param>
        /// <param name="propertyMeta">Mapped <c>PropertyMeta</c> object.</param>
        public ColumnInfo(object headerValue, PropertyMeta propertyMeta)
        {
            if (propertyMeta == null)
                throw new ArgumentNullException(nameof(propertyMeta));

            PropertyMeta = propertyMeta;
            HeaderValue = headerValue;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Refresh LastNonBlankValue property and get value according UseLastNonBlankValue property.
        /// </summary>
        /// <param name="value">The current cell value.</param>
        /// <returns>
        /// Same object as input parameter if UseLastNonBlankValue is false;
        /// otherwise return LastNonBlankValue.
        /// </returns>
        public object RefreshAndGetValue(object value)
        {
            // Specially check for string.
            if (string.IsNullOrWhiteSpace(value as string))
            {
                return PropertyMeta.UseLastNonBlankValue ? LastNonBlankValue : value;
            }

            LastNonBlankValue = value;

            return value;
        }

        #endregion
    }
}
