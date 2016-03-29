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
        public object HeaderValue { get; private set; }

        /// <summary>
        /// Column index.
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Mapped property for this column value according column name.
        /// </summary>
        public PropertyInfo Property { get; private set; }

        /// <summary>
        /// The column resolver.
        /// </summary>
        public ColumnResolver<TTarget> Resolver { get; set; }

        /// <summary>
        /// Indicate whether to use the last non-blank value.
        /// Typically handle the blank error in merged cells.
        /// </summary>
        public bool UseLastNonBlankValue { get; set; }

        /// <summary>
        /// The last non-blank value.
        /// </summary>
        public object LastNonBlankValue { get; set; }

        /// <summary>
        /// Refresh LastNonBlankValue property and get value according UseLastNonBlankValue property.
        /// </summary>
        /// <param name="value">New value.</param>
        /// <returns>
        /// Same object as input parameter if UseLastNonBlankValue is false;
        /// otherwise return LastNonBlankValue.
        /// </returns>
        public object RefreshAndGetValue(object value)
        {
            // Specially check for string.
            if (string.IsNullOrWhiteSpace(value as string))
            {
                return UseLastNonBlankValue ? LastNonBlankValue : value;
            }

            LastNonBlankValue = value;

            return value;
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo{TTarget}"/> class.
        /// </summary>
        /// <param name="headerValue">The header value</param>
        /// <param name="index">Column index</param>
        /// <param name="property">The target property info.</param>
        /// <param name="isItemOfCollection">whether the column should be treated as item(s) of a collection property.</param>
        public ColumnInfo(object headerValue, int index, PropertyInfo property, bool isItemOfCollection = false)
        {
            HeaderValue = headerValue;
            Index = index;
            Property = property;
        }

        #endregion
    }
}
