using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using System.Linq;

namespace Npoi.Mapper
{
    /// <summary>
    /// Provide static supportive functionalities for <see cref="Mapper"/> class.
    /// </summary>
    public static class MapHelper
    {
        #region Fields

        // Default chars that will be removed when mapping by column header name.
        private static readonly char[] DefaultIgnoredChars =
        {'`', '~', '!', '@', '#', '$', '%', '^', '&', '*', '-', '_', '+', '=', '|', ',', '.', '/', '?'};

        // Default chars to truncate column header name during mapping.
        private static readonly char[] DefaultTruncateChars = { '[', '<', '(', '{' };

        // Binding flags to lookup object properties.
        public const BindingFlags BindingFlag = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance;

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

        /// <summary>
        /// Cache for type of string during parsing.
        /// </summary>
        private static readonly Type StringType = typeof(string);

        /// <summary>
        /// Cache for type of DateTime during parsing.
        /// </summary>
        private static readonly Type DateTimeType = typeof(DateTime);

        #endregion

        #region Public Methods

        /// <summary>
        /// Load attributes to a dictionary.
        /// </summary>
        /// <typeparam name="T">The target type.</typeparam>
        /// <param name="attributes">Container to hold loaded attributes.</param>
        public static void LoadAttributes<T>(Dictionary<PropertyInfo, ColumnAttribute> attributes)
        {
            var type = typeof(T);

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var columnMeta = pi.GetCustomAttribute<ColumnAttribute>();
                var ignore = Attribute.IsDefined(pi, typeof(IgnoreAttribute));
                var useLastNonBlank = Attribute.IsDefined(pi, typeof(UseLastNonBlankValueAttribute));

                if (columnMeta == null && !ignore && !useLastNonBlank) continue;

                if (columnMeta == null) columnMeta = new ColumnAttribute
                {
                    Ignored = ignore ? new bool?(true) : null,
                    UseLastNonBlankValue = useLastNonBlank ? new bool?(true) : null
                };

                columnMeta.Property = pi;

                // Note that attribute from Map method takes precedence over Attribute meta data.
                columnMeta.MergeTo(attributes, false);
            }
        }

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
        /// Load cell data format by a specified row.
        /// </summary>
        /// <param name="dataRow">The row to load format from.</param>
        /// <param name="columns">The column collection to load formats into.</param>
        /// <param name="defaultFormats">The default formats specified for certain types.</param>
        public static void LoadDataFormats(IRow dataRow, IEnumerable<IColumnInfo> columns, Dictionary<Type, string> defaultFormats)
        {
            if (columns == null) return;

            foreach (var column in columns)
            {
                var pi = column.Attribute.Property;
                var type = pi?.PropertyType;

                if (column.Attribute.CustomFormat == null)
                {
                    if (type != null && !defaultFormats.ContainsKey(type))
                    {
                        type = column.Attribute.PropertyUnderlyingType;
                    }

                    if (type != null && defaultFormats.ContainsKey(type))
                    {
                        column.Attribute.CustomFormat = defaultFormats[type];
                    }
                }

                var cell = dataRow?.GetCell(column.Attribute.Index);

                if (cell != null)
                {
                    column.DataFormat = cell.CellStyle.DataFormat;
                }
            }
        }

        /// <summary>
        /// Get the cell style.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="customFormat">The custom format string.</param>
        /// <param name="builtinFormat">The built-in format number.</param>
        /// <param name="columnFormat">The default column format number.</param>
        /// <returns><c>ICellStyle</c> object for the given cell.</returns>
        public static ICellStyle GetCellStyle(ICell cell, string customFormat, short builtinFormat, short? columnFormat)
        {
            ICellStyle style = null;
            var workbook = cell?.Row.Sheet.Workbook;

            if (!string.IsNullOrWhiteSpace(customFormat))
            {
                if (CustomStyles.ContainsKey(customFormat))
                {
                    style = CustomStyles[customFormat];
                }
                else if (workbook != null)
                {
                    style = CreateCellStyle(workbook, customFormat);
                    CustomStyles[customFormat] = style;
                }
            }
            else if (workbook != null)
            {
                var format = builtinFormat != 0 ? builtinFormat : columnFormat ?? 0; /*default to 0*/

                if (format == 0)
                {
                    return null;
                }

                if (BuiltinStyles.ContainsKey(format))
                {
                    style = BuiltinStyles[format];
                }
                else
                {
                    style = CreateCellStyle(workbook, format);
                    BuiltinStyles[format] = style;
                }
            }

            return style;
        }

        /// <summary>
        /// Creates a <see cref="ICellStyle"/> object by the given <see cref="IWorkbook"/>.
        /// </summary>
        /// <param name="workbook">The <see cref="IWorkbook"/> object.</param>
        /// <param name="format">The custom format.</param>
        /// <returns>The <see cref="ICellStyle"/> object.</returns>
        public static ICellStyle CreateCellStyle(IWorkbook workbook, string format)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(format)) throw new ArgumentException($"Parameter '{nameof(format)}' cannot be null or white string.");

            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat(format);

            return style;
        }

        /// <summary>
        /// Creates a <see cref="ICellStyle"/> object by the given <see cref="IWorkbook"/>.
        /// </summary>
        /// <param name="workbook">The <see cref="IWorkbook"/> object.</param>
        /// <param name="format">The builtin format.</param>
        /// <returns>The <see cref="ICellStyle"/> object.</returns>
        public static ICellStyle CreateCellStyle(IWorkbook workbook, short format)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (format == 0) return null;

            var style = workbook.CreateCellStyle();
            style.DataFormat = format;

            return style;
        }

        /// <summary>
        /// Gets a <see cref="ICellStyle"/> object based on value's type by looking up the default format dictionary.
        /// </summary>
        /// <param name="workbook">The <see cref="IWorkbook"/> object.</param>
        /// <param name="value">The value object.</param>
        /// <param name="defaultFormats">Default format dictionary.</param>
        /// <returns>The <see cref="ICellStyle"/> object.</returns>
        public static ICellStyle GetDefaultStyle(IWorkbook workbook, object value, Dictionary<Type, string> defaultFormats)
        {
            if (value == null || workbook == null || defaultFormats == null) return null;

            ICellStyle style;
            var type = value.GetType();

            if (!defaultFormats.ContainsKey(type)) return null;

            var format = defaultFormats[type];

            if (string.IsNullOrWhiteSpace(format))
            {
                return null;
            }

            if (!CustomStyles.ContainsKey(format))
            {
                style = CreateCellStyle(workbook, format);
                CustomStyles[format] = style;
            }
            else
            {
                style = CustomStyles[format];
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

        /// <summary>
        /// Get mapped <c>PropertyInfo</c> by property selector expression.
        /// </summary>
        /// <typeparam name="T">The object type that property belongs to.</typeparam>
        /// <param name="propertySelector">The property selector expression.</param>
        /// <returns>The mapped <c>PropertyInfo</c> object.</returns>
        public static PropertyInfo GetPropertyInfoByExpression<T>(Expression<Func<T, object>> propertySelector)
        {
            var expression = propertySelector as LambdaExpression;

            if (expression == null)
                throw new ArgumentException("Only LambdaExpression is allowed!", nameof(propertySelector));

            var body = expression.Body.NodeType == ExpressionType.MemberAccess ?
                (MemberExpression)expression.Body :
                (MemberExpression)((UnaryExpression)expression.Body).Operand;

            // body.Member will return the MemberInfo of base class, so we have to get it from T...
            //return (PropertyInfo)body.Member;
            return typeof(T).GetMember(body.Member.Name)[0] as PropertyInfo;
        }

        /// <summary>
        /// Get refined name by removing specified chars and truncating by specified chars.
        /// </summary>
        /// <param name="name">The name to be refined.</param>
        /// <param name="ignoringChars">Chars will be removed from the name string.</param>
        /// <param name="truncatingChars">Chars used truncate the name string.</param>
        /// <returns>Refined name string.</returns>
        public static string GetRefinedName(string name, char[] ignoringChars, char[] truncatingChars)
        {
            if (name == null) return null;

            name = Regex.Replace(name, @"\s", "");
            var ignoredChars = ignoringChars ?? DefaultIgnoredChars;
            var truncateChars = truncatingChars ?? DefaultTruncateChars;

            name = ignoredChars.Aggregate(name, (current, c) => current.Replace(c, '\0'));

            var index = name.IndexOfAny(truncateChars);
            if (index >= 0) name = name.Remove(index);

            return name;
        }

        #endregion

        internal static void EnsureDefaultFormats(Dictionary<Type, string> defaultFormats)
        {
            //
            // For now, only take care DateTime.
            //

            if (!defaultFormats.ContainsKey(DateTimeType))
            {
                defaultFormats[DateTimeType] = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            }
        }

        // Convert the input object as the target type.
        internal static object ConvertType(object value, IColumnInfo column)
        {
            if (value == null || column == null || column.Attribute.Property == null) return null;

            var stringValue = value as string;
            var targeType = column.Attribute.Property.PropertyType;
            var underlyingType = column.Attribute.PropertyUnderlyingType;
            targeType = underlyingType ?? targeType;

            if (stringValue != null)
            {
                if (targeType == StringType)
                {
                    return stringValue;
                }

                if (targeType == DateTimeType)
                {
                    DateTime dateTime;
                    if (DateTime.TryParseExact(
                        stringValue,
                        column.Attribute.CustomFormat,
                        CultureInfo.CurrentCulture,
                        DateTimeStyles.AllowWhiteSpaces,
                        out dateTime)
                    )
                    {
                        return dateTime;
                    }
                }

                // Ensure we are not throwing exception and just read a null for nullable property.
                if (underlyingType != null)
                {
                    var converter = column.Attribute.PropertyUnderlyingConverter;
                    if (!converter.IsValid(value))
                    {
                        return null;
                    }
                }
            }

            return Convert.ChangeType(value, targeType);
        }
    }
}
