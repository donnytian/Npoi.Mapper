using System;
using System.Collections.Generic;
using System.Globalization;
using NPOI.SS.UserModel;

namespace Npoi.Mapper
{
    /// <summary>
    /// Provide static supportive functionalities for <see cref="Mapper"/> class.
    /// </summary>
    public static class MapHelper
    {
        #region Fields

        /// <summary>
        /// Collection of numeric types.
        /// </summary>
        private static readonly List<Type> NumericTypes = new List<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

        /// <summary>
        /// Store cached built-in styles to avoid create new ICellStyle for each cell.
        /// </summary>
        private static readonly Dictionary<short, ICellStyle> BuiltinStyles = new Dictionary<short, ICellStyle>();

        /// <summary>
        /// Store cached custom styles to avoid create new ICellStyle for each customized cell.
        /// </summary>
        private static readonly Dictionary<string, ICellStyle> CustomStyles = new Dictionary<string, ICellStyle>();

        #endregion

        #region Public Methods

        /// <summary>
        /// Extension for <see cref="IEnumerable{T}"/> object to handle each item.
        /// </summary>
        /// <typeparam name="T">The item type.</typeparam>
        /// <param name="sequence">The enumerable sequence.</param>
        /// <param name="action">Action to apply to each item.</param>
        public static void ForEach<T>(this IEnumerable<T> sequence, Action<T> action)
        {
            if (sequence == null) return;

            foreach (var item in sequence)
            {
                action(item);
            }
        }

        /// <summary>
        /// Clear cached data for cell styles and tracked column info.
        /// </summary>
        public static void ClearCache()
        {
            BuiltinStyles.Clear();
            CustomStyles.Clear();
        }

        /// <summary>
        /// Check if the given type is a numeric type.
        /// </summary>
        /// <param name="type">The type to be checked.</param>
        /// <returns><c>true</c> if it's numeric; otherwise <c>false</c>.</returns>
        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type);
        }

        /// <summary>
        /// Get the cell style.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="customFormat">The custom format string.</param>
        /// <param name="builtinFormat">The built-in format number.</param>
        /// <param name="defaultFormat">The default format number.</param>
        /// <returns><c>ICellStyle</c> object for the given cell.</returns>
        public static ICellStyle GetCellStyle(ICell cell, string customFormat, short builtinFormat, short defaultFormat = 0)
        {
            ICellStyle style = null;
            var workbook = cell?.Row.Sheet.Workbook;

            if (customFormat != null)
            {
                if (CustomStyles.ContainsKey(customFormat))
                {
                    style = CustomStyles[customFormat];
                }
                else if (workbook != null)
                {
                    style = workbook.CreateCellStyle();
                    style.DataFormat = workbook.CreateDataFormat().GetFormat(customFormat);
                    CustomStyles[customFormat] = style;
                }
            }
            else if (workbook != null)
            {
                var format = builtinFormat != 0 ? builtinFormat : defaultFormat;

                if (BuiltinStyles.ContainsKey(format))
                {
                    style = BuiltinStyles[format];
                }
                else
                {
                    style = workbook.CreateCellStyle();
                    style.DataFormat = format;
                    BuiltinStyles[format] = style;
                }
            }

            return style;
        }

        /// <summary>
        /// Get underline cell type if the cell is in formula.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>The underline cell type.</returns>
        public static CellType GetCellType(ICell cell)
        {
            return cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        }

        /// <summary>
        /// Try get cell value.
        /// </summary>
        /// <param name="cell">The cell to retrieve value.</param>
        /// <param name="targetType">Type of target property.</param>
        /// <param name="value">The returned value for cell.</param>
        /// <returns><c>true</c> if get value successfully; otherwise false.</returns>
        public static bool TryGetCellValue(ICell cell, Type targetType, out object value)
        {
            value = null;
            if (cell == null) return true;

            var success = true;

            switch (GetCellType(cell))
            {
                case CellType.String:

                    if (targetType?.IsEnum == true) // Enum type.
                    {
                        value = Enum.Parse(targetType, cell.StringCellValue, true);
                    }
                    else // String type.
                    {
                        value = cell.StringCellValue;
                    }

                    break;

                case CellType.Numeric:

                    if (DateUtil.IsCellDateFormatted(cell) || targetType == typeof(DateTime)) // DateTime type.
                    {
                        value = cell.DateCellValue;
                    }
                    else if (targetType?.IsEnum == true) // Enum type.
                    {
                        value = Enum.Parse(targetType, cell.NumericCellValue.ToString(CultureInfo.InvariantCulture));
                    }
                    else // Number type
                    {
                        value = cell.NumericCellValue;
                    }

                    break;

                case CellType.Boolean:

                    value = cell.BooleanCellValue;
                    break;

                case CellType.Error:
                case CellType.Unknown:
                case CellType.Blank:
                    // Dose nothing to keep return value null.
                    break;

                default:

                    success = false;

                    break;
            }

            return success;
        }

        #endregion
    }
}
