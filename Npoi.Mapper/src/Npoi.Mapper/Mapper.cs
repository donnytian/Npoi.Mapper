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

        // PropertyInfo map to PropertyMeta
        private Dictionary<PropertyInfo, PropertyMeta> MetaDict { get; } = new Dictionary<PropertyInfo, PropertyMeta>();

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
            var pi = GetPropertyInfoByExpression(propertySelector);
            var mapping = MetaDict.ContainsKey(pi)
                ? MetaDict[pi]
                : MetaDict[pi] = new PropertyMeta(columnName, pi, resolverType);

            mapping.ResolverType = resolverType;
            mapping.Ignored = false;
            mapping.Mapped = true;

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
            var pi = GetPropertyInfoByExpression(propertySelector);
            var mapping = MetaDict.ContainsKey(pi)
                ? MetaDict[pi]
                : MetaDict[pi] = new PropertyMeta(columnIndex, pi, resolverType);

            mapping.ResolverType = resolverType;
            mapping.Ignored = false;
            mapping.Mapped = true;

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
            var mapping = MetaDict.ContainsKey(pi)
                ? MetaDict[pi]
                : MetaDict[pi] = new PropertyMeta(null, pi);

            mapping.UseLastNonBlankValue = true;

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
            var mapping = MetaDict.ContainsKey(pi)
                ? MetaDict[pi]
                : MetaDict[pi] = new PropertyMeta(null, pi);

            mapping.Ignored = true;

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

        private void PrepareHeaders<T>(IRow headerRow, ICollection<ColumnInfo<T>> columns)
        {
            //
            // Column mapping priority:
            // Map<T> > ColumnAttribute > naming convention > MultiColumnsContainerAttribute > DefaultResolverType.
            //

            // Prepare a list of ColumnInfo.
            foreach (ICell header in headerRow)
            {
                // Custom mappings via Map<T> function.
                var column = GetColumnInfoByMappings<T>(header, MetaDict);

                // ColumnAttribute
                if (column == null)
                {
                    column = GetColumnInfoByColumnAttribute<T>(header);
                }

                // Naming convention.
                if (column == null && header.CellType == CellType.String)
                {
                    var s = header.StringCellValue;

                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        column = GetColumnInfoByName<T>(s.Trim(), header.ColumnIndex);
                    }
                }

                // MultiColumnsContainerAttribute
                if (column == null)
                {
                    column = GetColumnInfoByMultiColumnsContainerAttribute<T>(header);
                }

                // DefaultResolverType
                if (column == null)
                {
                    column = GetColumnInfoByResolverType<T>(header, DefaultResolverType);
                }

                if (column != null)
                {
                    var meta = column.PropertyMeta;
                    if (meta.Property != null && !meta.UseLastNonBlankValue)
                    {
                        meta.UseLastNonBlankValue = meta.Property
                            .GetCustomAttributes<UseLastNonBlankValueAttribute>().Any();
                    }

                    UpdateMapping(column);
                    columns.Add(column);
                }
            }
        }

        private ColumnInfo<T> GetColumnInfoByName<T>(string name, int index)
        {
            var type = typeof(T);

            // First attempt: search by exact string.
            var pi = type.GetProperty(name, BindingFlag);
            if (pi != null) return new ColumnInfo<T>(name, index, pi);

            // Second attempt: search display name of DisplayAttribute if any.
            foreach (var propertyInfo in GetProperties(type))
            {
                var atts = propertyInfo.GetCustomAttributes<DisplayAttribute>();

                if (atts.Any(att => string.Equals(@att.Name, @name, StringComparison.InvariantCultureIgnoreCase)))
                {
                    return new ColumnInfo<T>(name, index, propertyInfo);
                }
            }

            // Third attempt: remove space chars, '-', '_' and truncate by parentheses.
            name = Regex.Replace(name, @"\s", "").Replace("-", "").Replace("_", "");
            var bracketIndex = name.IndexOfAny(new[] { '(', '[', '{' });
            if (bracketIndex > 0) name = name.Remove(bracketIndex);
            pi = type.GetProperty(name, BindingFlag);

            return pi == null ? null : new ColumnInfo<T>(name, index, pi);
        }

        private static ColumnInfo<T> GetColumnInfoByMappings<T>(ICell header, Dictionary<PropertyInfo, PropertyMeta> mappings)
        {
            var type = typeof(T);
            var cellType = GetCellType(header);

            foreach (var pair in mappings)
            {
                if (pair.Key.ReflectedType != type || !pair.Value.Mapped || pair.Value.Ignored) continue;

                var mapping = pair.Value;

                if ((cellType == CellType.String && string.Equals(mapping.ColumnName, header.StringCellValue, StringComparison.CurrentCultureIgnoreCase))
                    || mapping.ColumnIndex == header.ColumnIndex)
                {
                    var resolver = pair.Value.ResolverType == null ?
                        null :
                        Activator.CreateInstance(pair.Value.ResolverType) as ColumnResolver<T>;

                    return new ColumnInfo<T>(GetHeaderValue(header), header.ColumnIndex, pair.Key)
                    {
                        Resolver = resolver
                    };
                }
            }

            return null;
        }

        private ColumnInfo<T> GetColumnInfoByColumnAttribute<T>(ICell header)
        {
            if (GetCellType(header) != CellType.String) return null;

            var type = typeof(T);

            foreach (var pi in GetProperties(type))
            {
                var att = pi.GetCustomAttributes<ColumnAttribute>().FirstOrDefault();

                if (att == null) continue;

                if (string.Equals(att.Name, header.StringCellValue, StringComparison.CurrentCultureIgnoreCase)
                    || att.Index == header.ColumnIndex)
                {
                    var resolver = att.ColumnResolverType == null ?
                        null :
                        Activator.CreateInstance(att.ColumnResolverType) as ColumnResolver<T>;

                    return new ColumnInfo<T>(header.StringCellValue, header.ColumnIndex, pi)
                    {
                        Resolver = resolver
                    };
                }
            }

            return null;
        }

        private ColumnInfo<T> GetColumnInfoByMultiColumnsContainerAttribute<T>(ICell header)
        {
            var type = typeof(T);

            foreach (var pi in GetProperties(type))
            {
                var att = pi.GetCustomAttributes<MultiColumnContainerAttribute>().FirstOrDefault();

                if (att == null) continue;

                var resolver = Activator.CreateInstance(att.ColumnResolverType) as ColumnResolver<T>;

                if (resolver == null) continue;

                var headerValue = GetHeaderValue(header);
                if (!resolver.TryResolveHeader(ref headerValue, header.ColumnIndex)) continue;

                return new ColumnInfo<T>(headerValue, header.ColumnIndex, pi)
                {
                    Resolver = resolver
                };
            }

            return null;
        }

        private static ColumnInfo<T> GetColumnInfoByResolverType<T>(ICell header, Type resolverType)
        {
            if (resolverType == null) return null;

            var resolver = Activator.CreateInstance(resolverType) as ColumnResolver<T>;

            if (resolver == null) return null;

            var headerValue = GetHeaderValue(header);

            if (!resolver.TryResolveHeader(ref headerValue, header.ColumnIndex)) return null;

            return new ColumnInfo<T>(headerValue, header.ColumnIndex, null)
            {
                Resolver = resolver
            };
        }

        private ColumnInfo<T> GetColumnInfo<T>(object headerValue, int index, PropertyInfo pi)
        {
            PropertyMeta pm;
            if (pi != null && MetaDict.ContainsKey(pi))
            {
                pm = MetaDict[pi];
            }
            else
            {
                pm = new PropertyMeta(index, pi);
            }

            return new ColumnInfo<T>(headerValue, pm);
        }

        private static RowInfo<T> GetRowData<T>(IEnumerable<ColumnInfo<T>> columns, IRow row, Func<T> objectInitializer)
        {
            var obj = objectInitializer == null ? Activator.CreateInstance<T>() : objectInitializer();
            var errorIndex = -1;
            var errorMessage = string.Empty;

            foreach (var column in columns)
            {
                try
                {
                    var cell = row.GetCell(column.PropertyMeta.ColumnIndex);
                    var propertyType = column.PropertyMeta.Property?.PropertyType;
                    object valueObj;

                    if (!TryGetCellValue(cell, propertyType, out valueObj))
                    {
                        errorIndex = column.PropertyMeta.ColumnIndex;
                        errorMessage = "CellType is not supported yet!";
                        break;
                    }

                    valueObj = column.RefreshAndGetValue(valueObj);

                    if (column.Resolver != null)
                    {
                        if (!column.Resolver.TryResolveCell(column, valueObj, obj))
                        {
                            errorIndex = column.PropertyMeta.ColumnIndex;
                            errorMessage = "Returned failure by custom cell resolver!";
                            break;
                        }
                    }
                    else if (propertyType != null && valueObj != null)
                    {
                        // Change types between IConvertible objects, such as double, float, int and etc.
                        var value = Convert.ChangeType(valueObj, propertyType);
                        column.PropertyMeta.Property.SetValue(obj, value);
                    }
                    else
                    {
                        // If we go this far, keep target property untouched...
                    }
                }
                catch (Exception e)
                {
                    errorIndex = column.PropertyMeta.ColumnIndex;
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

        private IEnumerable<PropertyInfo> GetProperties(Type type)
        {
            if (type == null) yield break;

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                if (MetaDict.ContainsKey(pi) && MetaDict[pi].Ignored)
                {
                    continue;
                }

                yield return pi;
            }
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

        private void UpdateMapping<T>(ColumnInfo<T> column)
        {
            if (column.PropertyMeta.Property == null) return;
            if (!MetaDict.ContainsKey(column.PropertyMeta.Property)) return;

            var mapping = MetaDict[column.PropertyMeta.Property];
            column.PropertyMeta.UseLastNonBlankValue = mapping.UseLastNonBlankValue;
        }

        #endregion
    }
}
