using System;
using System.Collections.Generic;
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
    public class MapHelper
    {
        #region Fields

        /// <summary>
        /// Stores cached built-in styles to avoid create new ICellStyle for each cell.
        /// </summary>
        private readonly Dictionary<short, ICellStyle> _builtinStyles = new Dictionary<short, ICellStyle>();

        /// <summary>
        /// Stores cached custom styles to avoid create new ICellStyle for each customized cell.
        /// </summary>
        private readonly Dictionary<string, ICellStyle> _customStyles = new Dictionary<string, ICellStyle>();

        // Column chars that will be used for Excel columns.
        // e.g. Column A is the first column, Column AA is the 27th column.
        private static readonly char[] ColumnChars =
        {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'};

        // Default chars that will be removed when mapping by column header name.
        private static readonly char[] DefaultIgnoredChars =
        {'`', '~', '!', '@', '#', '$', '%', '^', '&', '*', '-', '_', '+', '=', '|', ',', '.', '/', '?'};

        // Default chars to truncate column header name during mapping.
        private static readonly char[] DefaultTruncateChars = { '[', '<', '(', '{' };

        // Binding flags to lookup object properties.
        public const BindingFlags BindingFlag = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance;


        /// <summary>
        /// Caches for type of string during parsing.
        /// </summary>
        public static readonly Type StringType = typeof(string);

        /// <summary>
        /// Caches for type of DateTime during parsing.
        /// </summary>
        public static readonly Type DateTimeType = typeof(DateTime);

        /// <summary>
        /// Caches for type of object.
        /// </summary>
        public static readonly Type ObjectType = typeof(object);

        public static readonly Type GuidType = typeof(Guid);

        /// <summary>
        /// The maximum row number during the detection for column type and style.
        /// </summary>
        public int MaxLookupRowNum { get; set; } = 20;

        #endregion

        #region Public Methods

        /// <summary>
        /// Loads attributes to a dictionary.
        /// </summary>
        /// <param name="attributes">Container to hold loaded attributes.</param>
        /// <param name="type">The type object.</param>
        public static void LoadAttributes(Dictionary<PropertyInfo, ColumnAttribute> attributes, Type type)
        {
            if (type == null) return;

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var columnMeta = pi.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
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
        /// Loads dynamic attributes to a dictionary.
        /// </summary>
        /// <param name="attributes">Container to hold loaded attributes.</param>
        /// <param name="dynamicAttributes">Container for dynamic attributes to be loaded.</param>
        /// <param name="dynamicType">The type object created by runtime.</param>
        public static void LoadDynamicAttributes(Dictionary<PropertyInfo, ColumnAttribute> attributes, Dictionary<string, ColumnAttribute> dynamicAttributes, Type dynamicType)
        {
            foreach (var pair in dynamicAttributes)
            {
                var pi = dynamicType.GetProperty(pair.Key);

                if (pi != null)
                {
                    pair.Value.Property = pi;
                    pair.Value.MergeTo(attributes);
                }
            }
        }

        /// <summary>
        /// Clear cached data for cell styles and tracked column info.
        /// </summary>
        public void ClearCache()
        {
            _builtinStyles.Clear();
            _customStyles.Clear();
        }

        /// <summary>
        /// Load cell data format by a specified row.
        /// </summary>
        /// <param name="sheet">The sheet to load format from.</param>
        /// <param name="firstDataRowIndex">The index for the first row to detect.</param>
        /// <param name="columns">The column collection to load formats into.</param>
        /// <param name="defaultFormats">The default formats specified for certain types.</param>
        public void LoadDataFormats(ISheet sheet, int firstDataRowIndex, IEnumerable<IColumnInfo> columns, Dictionary<Type, string> defaultFormats)
        {
            if (sheet == null) return;
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

                var rowIndex = firstDataRowIndex >= 0 ? firstDataRowIndex : sheet.FirstRowNum + 1;
                while (rowIndex <= sheet.LastRowNum && rowIndex <= MaxLookupRowNum)
                {
                    var dataRow = sheet.GetRow(rowIndex);
                    var cell = dataRow?.GetCell(column.Attribute.Index);

                    rowIndex++;
                    if (cell?.CellStyle == null) continue;
                    column.DataFormat = cell.CellStyle.DataFormat;
                    break;
                }
            }
        }

        /// <summary>
        /// Get the cell style.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="customFormat">The custom format string.</param>
        /// <param name="columnFormat">The default column format number.</param>
        /// <returns><c>ICellStyle</c> object for the given cell; null if not format specified.</returns>
        public ICellStyle GetCellStyle(ICell cell, string customFormat, short? columnFormat)
        {
            ICellStyle style = null;
            var workbook = cell?.Row.Sheet.Workbook;

            if (!string.IsNullOrWhiteSpace(customFormat))
            {
                if (_customStyles.ContainsKey(customFormat))
                {
                    style = _customStyles[customFormat];
                }
                else if (workbook != null)
                {
                    style = CreateCellStyle(workbook, customFormat);
                    _customStyles[customFormat] = style;
                }
            }
            else if (workbook != null)
            {
                var format = columnFormat ?? 0; // Defaults to 0.

                if (format == 0)
                {
                    return null;
                }

                if (_builtinStyles.ContainsKey(format))
                {
                    style = _builtinStyles[format];
                }
                else
                {
                    style = CreateCellStyle(workbook, format);
                    _builtinStyles[format] = style;
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
        public ICellStyle GetDefaultStyle(IWorkbook workbook, object value, Dictionary<Type, string> defaultFormats)
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

            if (!_customStyles.ContainsKey(format))
            {
                style = CreateCellStyle(workbook, format);
                _customStyles[format] = style;
            }
            else
            {
                style = _customStyles[format];
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

                    value = cell.StringCellValue;

                    break;

                case CellType.Numeric:

                    if (DateUtil.IsCellDateFormatted(cell) || targetType == typeof(DateTime)) // DateTime type.
                    {
                        value = cell.DateCellValue;
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

            name = Regex.Replace(name, @"\s", string.Empty);
            var ignoredChars = ignoringChars ?? DefaultIgnoredChars;
            var truncateChars = truncatingChars ?? DefaultTruncateChars;

            name = ignoredChars.Aggregate(name, (current, c) => current.Replace(c.ToString(), string.Empty));

            var index = name.IndexOfAny(truncateChars);
            if (index >= 0) name = name.Remove(index);

            return name;
        }

        /// <summary>
        /// Get a valid variable name by removing specified chars and truncating by specified chars.
        /// </summary>
        /// <param name="rawName">The name to be revised as a valid variable name.</param>
        /// <param name="ignoringChars">Chars will be removed from the name string.</param>
        /// <param name="truncatingChars">Chars used truncate the name string.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <returns>A valid variable name based on the rawName.</returns>
        public static string GetVariableName(string rawName, char[] ignoringChars, char[] truncatingChars, int columnIndex)
        {
            rawName = GetRefinedName(rawName, ignoringChars, truncatingChars);

            if (string.IsNullOrEmpty(rawName))
            {
                rawName = GetExcelColumnName(columnIndex);
            }

            return rawName;
        }

        /// <summary>
        /// Get column name in Excel, like A, B, AC.
        /// </summary>
        /// <param name="columnIndex">The column index.</param>
        /// <returns>The column name that represent the order in Excel.</returns>
        public static string GetExcelColumnName(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex > 16383) throw new ArgumentOutOfRangeException(nameof(columnIndex));

            var columnName = string.Empty;
            var result = columnIndex;
            do
            {
                var reminder = result % ColumnChars.Length;
                columnName = ColumnChars[reminder] + columnName;
                result = result / ColumnChars.Length - 1;
            } while (result != -1);

            return columnName;
        }

        /// <summary>
        /// Determines the data type for the specified column according the first non-blank data cell.
        /// </summary>
        /// <param name="sheet">The sheet contains data.</param>
        /// <param name="headerRowIndex">The row index for the header, pass -1 if no header.</param>
        /// <param name="columnIndex">The index for the column.</param>
        /// <returns>The type object.</returns>
        public Type InferColumnDataType(ISheet sheet, int headerRowIndex, int columnIndex)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));

            Type type = null;
            var rowIndex = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
            var typeDetected = false;

            while (!typeDetected && rowIndex <= sheet.LastRowNum && rowIndex <= MaxLookupRowNum)
            {
                var row = sheet.GetRow(rowIndex);

                var cell = row?.GetCell(columnIndex);
                if (cell != null)
                {
                    var cellType = GetCellType(cell);
                    typeDetected = true;
                    switch (cellType)
                    {
                        case CellType.Boolean:
                            type = typeof(bool);
                            break;
                        case CellType.Numeric:
                            type = DateUtil.IsCellDateFormatted(cell) ? DateTimeType : typeof(double);
                            break;
                        case CellType.String:
                            type = StringType;
                            break;
                        default:
                            typeDetected = false;
                            break;
                    }
                }
                rowIndex++;
            }

            return type;
        }

        #endregion

        internal static void EnsureDefaultFormats(IEnumerable<IColumnInfo> columns, Dictionary<Type, string> defaultFormats)
        {
            //
            // For now, only take care DateTime.
            //
            if (!defaultFormats.ContainsKey(DateTimeType))
            {
                defaultFormats[DateTimeType] = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            }

            foreach (var column in columns)
            {
                var attributes = column.Attribute;
                if (column.DataFormat == null &&
                    attributes.Property != null &&
                    attributes.CustomFormat == null)
                {
                    var type = attributes.Property.PropertyType;

                    if (defaultFormats.ContainsKey(type))
                    {
                        attributes.CustomFormat = defaultFormats[type];
                    }
                }
            }
        }

        // Try to convert the input object as the target type.
        internal static bool TryConvertType(object value, IColumnInfo column, out object result)
        {
            result = null;
            if (column == null || column.Attribute.Property == null) return false;
            if (value == null) return true;

            var stringValue = value as string;
            var targetType = column.Attribute.Property.PropertyType;
            var underlyingType = column.Attribute.PropertyUnderlyingType;
            targetType = underlyingType ?? targetType;

            if (stringValue != null)
            {
                if (targetType == StringType)
                {
                    result = stringValue;
                    return true;
                }

                if (targetType == DateTimeType)
                {
                    if (DateTime.TryParseExact(stringValue, column.Attribute.CustomFormat, CultureInfo.CurrentCulture,
                        DateTimeStyles.AllowWhiteSpaces, out DateTime dateTime)
                    )
                    {
                        result = dateTime;
                        return true;
                    }
                }

                if (targetType.IsNumeric() && double.TryParse(stringValue, NumberStyles.Any, null, out double doubleResult))
                {
                    result = Convert.ChangeType(doubleResult, targetType);
                    return true;
                }

                if (targetType.IsEnum)
                {
                    result = Enum.Parse(targetType, stringValue, true);
                    return true;
                }

                if (targetType == GuidType)
                {
                    var parsed = Guid.TryParse(stringValue, out var guidResult);
                    result = guidResult;
                    return parsed;
                }

                // Ensure we are not throwing exception and just read a null for nullable property.
                if (underlyingType != null)
                {
                    if (string.IsNullOrWhiteSpace(stringValue))
                    {
                        return true;
                    }

                    var converter = column.Attribute.PropertyUnderlyingConverter;
                    if (!converter.IsValid(value))
                    {
                        return false;
                    }
                }
            }

            try
            {
                result = Convert.ChangeType(value, targetType);
            }
            catch
            {
                return false;
            }

            return true;
        }

        // Gets the concrete type instead of the type of object.
        internal static Type GetConcreteType<T>(T[] objects)
        {
            var type = typeof(T);
            if (type != ObjectType) return type;

            foreach (var o in objects)
            {
                if (o == null) continue;
                type = o.GetType();
                if (type != ObjectType)
                {
                    break;
                }
            }

            return type;
        }
    }
}
