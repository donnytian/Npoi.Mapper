using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

using NPOI.SS.UserModel;
using Npoi.Mapper.Attributes;

namespace Npoi.Mapper
{
    /// <summary>
    /// Import Excel row data as object.
    /// </summary>
    public class Mapper
    {
        #region Fields

        private const BindingFlags BindingFlag = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance;

        #endregion

        #region Properties

        // PropertyInfo map to ColumnAttribute
        private Dictionary<PropertyInfo, ColumnAttribute> Attributes { get; } = new Dictionary<PropertyInfo, ColumnAttribute>();

        // Type of resolver to handle unrecognized columns.
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public Type DefaultResolverType { get; set; }

        // Excel file workbook.
        public IWorkbook Workbook { get; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class.
        /// </summary>
        /// <param name="stream">The input Excel(XLS, XLSX) file stream</param>
        public Mapper(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            using (stream)
            {
                Workbook = WorkbookFactory.Create(stream, ImportOption.SheetContentOnly);
            }
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class.
        /// </summary>
        /// <param name="workbook">The input IWorkbook object.</param>
        public Mapper(IWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Workbook = workbook;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class.
        /// </summary>
        /// <param name="filePath">The path of Excel file.</param>
        public Mapper(string filePath) : this(new FileStream(filePath, FileMode.Open))
        {
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Map property to a column by name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">The type of resolver.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Map<T>(string columnName, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            if (columnName == null)
                throw new ArgumentNullException(nameof(columnName));

            var pi = GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            var attribute = new ColumnAttribute()
            {
                Name = columnName,
                Property = pi,
                ResolverType = resolverType,
                Ignored = false
            };

            Attributes[pi] = Attributes.ContainsKey(pi)
                ? MergeAttribute(Attributes[pi], attribute)
                : Attributes[pi] = attribute;

            return this;
        }

        /// <summary>
        /// Map property to a column by index.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">The type of resolver.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Map<T>(int columnIndex, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            if (columnIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(columnIndex));

            var pi = GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            var attribute = new ColumnAttribute()
            {
                Index = columnIndex,
                Property = pi,
                ResolverType = resolverType,
                Ignored = false
            };

            Attributes[pi] = Attributes.ContainsKey(pi)
                ? MergeAttribute(Attributes[pi], attribute)
                : Attributes[pi] = attribute;

            return this;
        }

        /// <summary>
        /// Specify to use last non-blank value for a property.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper UseLastNonBlankValue<T>(Expression<Func<T, object>> propertySelector)
        {
            var pi = GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            var attribute = new ColumnAttribute()
            {
                Property = pi,
                UseLastNonBlankValue = true
            };

            Attributes[pi] = Attributes.ContainsKey(pi)
                ? MergeAttribute(Attributes[pi], attribute)
                : Attributes[pi] = attribute;

            return this;
        }

        /// <summary>
        /// Specify to ignore a property.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Ignore<T>(Expression<Func<T, object>> propertySelector)
        {
            var pi = GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            var attribute = new ColumnAttribute()
            {
                Property = pi,
                Ignored = true
            };

            Attributes[pi] = Attributes.ContainsKey(pi)
                ? MergeAttribute(Attributes[pi], attribute)
                : Attributes[pi] = attribute;

            return this;
        }

        /// <summary>
        /// Get objects of target type by converting rows in the sheet with specified index (zero based).
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="sheetIndex">The sheet index; default is 0.</param>
        /// <param name="maxErrorRows">The maximum error rows before stop reading; default is 10.</param>
        /// <param name="objectInitializer">Factory method to create a new target object.</param>
        /// <returns>Objects of target type</returns>
        public IEnumerable<RowInfo<T>> Take<T>(int sheetIndex = 0, int maxErrorRows = 10, Func<T> objectInitializer = null)
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return TakeByHeader(sheet, maxErrorRows, objectInitializer);
        }

        /// <summary>
        /// Get objects of target type by converting rows in the sheet with specified name.
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="maxErrorRows">The maximum error rows before stopping read; default is 10.</param>
        /// <param name="objectInitializer">Factory method to create a new target object.</param>
        /// <returns>Objects of target type</returns>
        public IEnumerable<RowInfo<T>> Take<T>(string sheetName, int maxErrorRows = 10, Func<T> objectInitializer = null)
        {
            var sheet = Workbook.GetSheet(sheetName);
            return TakeByHeader(sheet, maxErrorRows, objectInitializer);
        }

        #endregion

        #region Private Methods

        private IEnumerable<RowInfo<T>> TakeByHeader<T>(ISheet sheet, int maxErrorRows, Func<T> objectInitializer = null)
        {
            if (sheet == null || sheet.PhysicalNumberOfRows < 2)
            {
                yield break;
            }

            var headerIndex = sheet.FirstRowNum;
            var headerRow = sheet.GetRow(headerIndex);
            var headers = new List<ColumnInfo<T>>();

            ScanAttributes<T>(Attributes);
            PrepareHeaders(headerRow, headers);

            // Loop rows in file. Generate one target object for each row.
            var errorCount = 0;
            foreach (IRow row in sheet)
            {
                if (maxErrorRows > 0 && errorCount >= maxErrorRows) break;
                if (row.RowNum == headerIndex) continue;

                var data = GetRowData(headers, row, objectInitializer);

                if (data.ErrorColumnIndex >= 0) errorCount++;

                yield return data;
            }
        }

        private static void ScanAttributes<T>(Dictionary<PropertyInfo, ColumnAttribute> attributes)
        {
            if (attributes == null) return;

            var type = typeof(T);

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var attributeMeta = pi.GetCustomAttributes<ColumnAttribute>().FirstOrDefault();

                if (attributeMeta == null) continue;

                attributeMeta.Property = pi;

                if (attributes.ContainsKey(pi))
                {
                    // Fluent attribute takes precedence over attribute meta.
                    attributes[pi] = MergeAttribute(attributeMeta, attributes[pi]);
                }
                else
                {
                    attributes[pi] = attributeMeta;
                }
            }
        }

        private static ColumnAttribute MergeAttribute(ColumnAttribute oldAttribute, ColumnAttribute newAttribute)
        {
            //
            // New attribute takes precedence over old attribute.
            //

            if (newAttribute.Index < 0) newAttribute.Index = oldAttribute.Index;

            if (newAttribute.Name == null) newAttribute.Name = oldAttribute.Name;

            if (newAttribute.Property == null) newAttribute.Property = oldAttribute.Property;

            if (newAttribute.ResolverType == null) newAttribute.ResolverType = oldAttribute.ResolverType;

            if (!newAttribute.Ignored) newAttribute.Ignored = oldAttribute.Ignored;

            if (!newAttribute.UseLastNonBlankValue) newAttribute.UseLastNonBlankValue = oldAttribute.UseLastNonBlankValue;

            return newAttribute;
        }

        private void PrepareHeaders<T>(IRow headerRow, ICollection<ColumnInfo<T>> columns)
        {
            //
            // Column mapping priority:
            // Map<T> > ColumnAttribute > naming convention > MultiColumnsContainerAttribute > DefaultResolverType.
            //

            // Prepare a list of ColumnInfo.
            foreach (ICell header in headerRow)
            {
                // Custom mappings via attributes.
                var column = GetColumnInfoByAttribute<T>(header);

                // Naming convention.
                if (column == null && header.CellType == CellType.String)
                {
                    var s = header.StringCellValue;

                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        column = GetColumnInfoByName<T>(s.Trim(), header.ColumnIndex);
                    }
                }

                // DefaultResolverType
                if (column == null)
                {
                    column = GetColumnInfoByResolverType<T>(header, DefaultResolverType);
                }

                if (column != null)
                {
                    columns.Add(column);
                }
            }
        }

        private ColumnInfo<T> GetColumnInfoByAttribute<T>(ICell header)
        {
            var type = typeof(T);
            var cellType = GetCellType(header);
            var index = header.ColumnIndex;

            foreach (var pair in Attributes)
            {
                if (pair.Key.ReflectedType != type || pair.Value.Ignored) continue;

                var attribute = pair.Value;
                var indexMatch = attribute.Index == index;
                var nameMatch = cellType == CellType.String && string.Equals(attribute.Name, header.StringCellValue);

                if (indexMatch || nameMatch)
                {
                    attribute = attribute.Clone();
                    attribute.Index = index;

                    var resolver = pair.Value.ResolverType == null ?
                        null :
                        Activator.CreateInstance(pair.Value.ResolverType) as ColumnResolver<T>;

                    return new ColumnInfo<T>(GetHeaderValue(header), attribute)
                    {
                        Resolver = resolver
                    };
                }

                if (attribute.Index < 0 && attribute.Name == null && attribute.ResolverType != null)
                {
                    var resolver = Activator.CreateInstance(pair.Value.ResolverType) as ColumnResolver<T>;

                    if (resolver != null)
                    {
                        var headerValue = GetHeaderValue(header);
                        if (resolver.IsColumnMapped(ref headerValue, index))
                        {
                            attribute = attribute.Clone();
                            attribute.Index = index;
                            return new ColumnInfo<T>(headerValue, attribute)
                            {
                                Resolver = resolver
                            };
                        }
                    }
                }
            }

            return null;
        }

        private ColumnInfo<T> GetColumnInfoByName<T>(string name, int index)
        {
            var type = typeof(T);

            // First attempt: search by string (ignore case).
            var pi = type.GetProperty(name, BindingFlag);

            if (pi == null)
            {
                // Second attempt: search display name of DisplayAttribute if any.
                foreach (var propertyInfo in type.GetProperties(BindingFlag))
                {
                    var atts = propertyInfo.GetCustomAttributes<DisplayAttribute>();

                    if (atts.Any(att => string.Equals(@att.Name, @name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        pi = propertyInfo;
                        break;
                    }
                }
            }

            if (pi == null)
            {
                // Third attempt: remove space chars, '-', '_', ',', '.' and truncate by parentheses.
                name = Regex.Replace(name, @"\s", "").Replace("-", "").Replace("_", "").Replace(",", "").Replace("_", "");
                var bracketIndex = name.IndexOfAny(new[] { '(', '[', '{', '<' });
                if (bracketIndex > 0) name = name.Remove(bracketIndex);
                pi = type.GetProperty(name, BindingFlag);
            }

            ColumnAttribute attribute = null;

            if (pi != null && Attributes.ContainsKey(pi))
            {

                attribute = Attributes[pi].Clone();
                attribute.Index = index;
                if (attribute.Ignored) return null;
            }

            return pi == null ? null : attribute == null ? new ColumnInfo<T>(name, index, pi) : new ColumnInfo<T>(name, attribute);
        }

        private static ColumnInfo<T> GetColumnInfoByResolverType<T>(ICell header, Type resolverType)
        {
            if (resolverType == null) return null;

            var resolver = Activator.CreateInstance(resolverType) as ColumnResolver<T>;

            if (resolver == null) return null;

            var headerValue = GetHeaderValue(header);

            if (!resolver.IsColumnMapped(ref headerValue, header.ColumnIndex)) return null;

            return new ColumnInfo<T>(headerValue, header.ColumnIndex, null)
            {
                Resolver = resolver
            };
        }

        private static RowInfo<T> GetRowData<T>(IEnumerable<ColumnInfo<T>> columns, IRow row, Func<T> objectInitializer)
        {
            var obj = objectInitializer == null ? Activator.CreateInstance<T>() : objectInitializer();
            var errorIndex = -1;
            var errorMessage = string.Empty;

            foreach (var column in columns)
            {
                if (column.Attribute.Index < 0) continue;
                var index = column.Attribute.Index;

                try
                {
                    var cell = row.GetCell(index);
                    var propertyType = column.Attribute.Property?.PropertyType;
                    object valueObj;

                    if (!TryGetCellValue(cell, propertyType, out valueObj))
                    {
                        errorIndex = index;
                        errorMessage = "CellType is not supported yet!";
                        break;
                    }

                    valueObj = column.RefreshAndGetValue(valueObj);

                    if (column.Resolver != null)
                    {
                        if (!column.Resolver.TryResolveCell(column, valueObj, obj))
                        {
                            errorIndex = index;
                            errorMessage = "Returned failure by custom cell resolver!";
                            break;
                        }
                    }
                    else if (propertyType != null && valueObj != null)
                    {
                        // Change types between IConvertible objects, such as double, float, int and etc.
                        var value = Convert.ChangeType(valueObj, propertyType);
                        column.Attribute.Property.SetValue(obj, value);
                    }
                    else
                    {
                        // If we go this far, keep target property untouched...
                    }
                }
                catch (Exception e)
                {
                    errorIndex = index;
                    errorMessage = e.Message;
                    break;
                }
            }

            if (errorIndex >= 0) obj = default(T);

            return new RowInfo<T>(row.RowNum, obj, errorIndex, errorMessage);
        }

        private static object GetHeaderValue(ICell header)
        {
            object value;
            var cellType = header.CellType;

            if (cellType == CellType.Formula)
            {
                cellType = header.CachedFormulaResultType;
            }

            switch (cellType)
            {
                case CellType.Numeric:
                    value = header.NumericCellValue;
                    break;

                case CellType.String:
                    value = header.StringCellValue;
                    break;

                default:
                    value = null;
                    break;
            }

            return value;
        }

        private static bool TryGetCellValue(ICell cell, Type targetType, out object value)
        {
            value = null;
            if (cell == null) return true;

            var success = true;

            switch (GetCellType(cell))
            {
                case CellType.String:
                    if (targetType != null && targetType.IsEnum) // Enum type.
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
                    else // Number type
                    {
                        value = cell.NumericCellValue;
                    }

                    break;

                case CellType.Blank:
                    // Dose nothing to keep return value null.
                    break;

                default: // TODO. Support other types.
                    success = false;

                    break;
            }

            return success;
        }

        private static CellType GetCellType(ICell cell)
        {
            return cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        }

        private static PropertyInfo GetPropertyInfoByExpression<T>(Expression<Func<T, object>> propertySelector)
        {
            var expression = propertySelector as LambdaExpression;

            if (expression == null)
                throw new ArgumentException("Only LambdaExpression is allowed!", nameof(propertySelector));

            var body = expression.Body.NodeType == ExpressionType.MemberAccess ?
                (MemberExpression)expression.Body :
                (MemberExpression)((UnaryExpression)expression.Body).Operand;

            return (PropertyInfo)body.Member;
        }

        #endregion
    }
}
