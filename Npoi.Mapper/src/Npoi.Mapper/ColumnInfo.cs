using System;
using System.Reflection;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required for one column when mapping between object and file rows.
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public class ColumnInfo<TTarget>
    {
        #region Fields

        #endregion

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
        public ColumnAttribute Attribute { get; set; }

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
            Attribute = new ColumnAttribute()
            {
                Name = columnName,
                Property = pi
            };
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
            Attribute = new ColumnAttribute()
            {
                Index = columnIndex,
                Property = pi
            };
        }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo{TTarget}"/> class.
        /// </summary>
        /// <param name="headerValue">The header value</param>
        /// <param name="attribute">Mapped <c>PropertyMeta</c> object.</param>
        public ColumnInfo(object headerValue, ColumnAttribute attribute)
        {
            if (attribute == null)
                throw new ArgumentNullException(nameof(attribute));

            Attribute = attribute;
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
                return Attribute.UseLastNonBlankValue == true ? LastNonBlankValue : value;
            }

            LastNonBlankValue = value;

            return value;
        }

        /// <summary>
        /// Set style for the cell.
        /// </summary>
        /// <param name="cell">The cell to be set.</param>
        /// <param name="defaultFormat">The default format.</param>
        public void SetCellStyle(ICell cell, short defaultFormat = 0)
        {
            if (cell != null)
            {
                cell.CellStyle = MapHelper.GetCellStyle(cell, Attribute.CustomFormat, Attribute.BuiltinFormat, defaultFormat);
            }
        }

        #endregion
    }
}
