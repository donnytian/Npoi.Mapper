using System;
using System.Collections.Generic;
using System.Reflection;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required for one column when mapping between object and file rows.
    /// </summary>
    public class ColumnInfo : IColumnInfo
    {
        #region Fields

        // For cache purpose, avoid lookup style dictionary for every cell.
        private ICellStyle _headerStyle;
        private ICellStyle _dataStyle;
        private bool _headerStyleCached;
        private bool _dataStyleCached;

        #endregion

        #region Properties

        /// <summary>
        /// Value for the column header.
        /// </summary>
        public object HeaderValue { get; set; }

        /// <summary>
        /// The mapped property information.
        /// </summary>
        public ColumnAttribute Attribute { get; set; }

        /// <summary>
        /// The last non-blank cell value.
        /// </summary>
        public object LastNonBlankValue { get; set; }

        /// <summary>
        /// The current cell value, might be used for custom resolving.
        /// </summary>
        public object CurrentValue { get; set; }

        /// <summary>
        /// Get or set the header cell format.
        /// </summary>
        public short? HeaderFormat { get; set; }

        /// <summary>
        /// Get or set the data cell format.
        /// </summary>
        public short? DataFormat { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnInfo"/> class.
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
        /// Initialize a new instance of <see cref="ColumnInfo"/> class.
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
        /// Initialize a new instance of <see cref="ColumnInfo"/> class.
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
        /// Refresh LastNonBlankValue and CurrentValue property then return value according UseLastNonBlankValue property.
        /// </summary>
        /// <param name="value">The current cell value.</param>
        /// <returns>
        /// Same object as input parameter if UseLastNonBlankValue is false;
        /// otherwise return LastNonBlankValue.
        /// </returns>
        public object RefreshAndGetValue(object value)
        {
            CurrentValue = value;

            // Specially check for string.
            if (string.IsNullOrWhiteSpace(value as string))
            {
                return Attribute.UseLastNonBlankValue == true ? LastNonBlankValue : value;
            }

            LastNonBlankValue = value;

            return value;
        }

        /// <summary>
        /// Set style of the cell for export.
        /// Assume the cell belongs to current column.
        /// </summary>
        /// <param name="cell">The cell to be set.</param>
        /// <param name="value">The cell value object.</param>
        /// <param name="isHeader">If <c>true</c>, use HeaderFormat; otherwise use DataFormat.</param>
        /// <param name="defaultFormats">The default formats dictionary.</param>
        /// <param name="helper">The helper object.</param>
        public void SetCellStyle(ICell cell, object value, bool isHeader, Dictionary<Type, string> defaultFormats, MapHelper helper)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            if (isHeader && !_headerStyleCached)
            {
                _headerStyle = helper.GetCellStyle(cell, null, HeaderFormat);

                if (_headerStyle == null && HeaderValue != null)
                {
                    _headerStyle = helper.GetDefaultStyle(cell.Sheet.Workbook, HeaderValue, defaultFormats);
                }

                _headerStyleCached = true;
            }
            else if (!isHeader && !_dataStyleCached)
            {
                _dataStyle = helper.GetCellStyle(cell, Attribute.CustomFormat, DataFormat);

                if (_dataStyle == null && value != null)
                {
                    _dataStyle = helper.GetDefaultStyle(cell.Sheet.Workbook, value, defaultFormats);
                }

                _dataStyleCached = true;
            }

            cell.CellStyle = isHeader ? _headerStyle : _dataStyle;
        }

        #endregion
    }
}
