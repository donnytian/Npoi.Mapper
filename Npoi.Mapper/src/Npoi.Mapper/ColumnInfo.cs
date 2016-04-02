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
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public class ColumnInfo<TTarget>
    {
        private static readonly List<Type> NumericTypes = new List<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

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
                return Attribute.UseLastNonBlankValue ? LastNonBlankValue : value;
            }

            LastNonBlankValue = value;

            return value;
        }

        public void SetCellFormat(ICell cell, short defaultFormat = 0)
        {
            var workbook = cell.Row.Sheet.Workbook;
            var style = workbook.CreateCellStyle();
            style.DataFormat = Attribute.CustomFormat != null 
                ? workbook.CreateDataFormat().GetFormat(Attribute.CustomFormat)
                : Attribute.BuiltinFormat != 0 ? Attribute.BuiltinFormat : defaultFormat;
            cell.CellStyle = style;
        }

        public bool IsNumeric(Type type)
        {
            return NumericTypes.Contains(type);
        }

        #endregion
    }
}
