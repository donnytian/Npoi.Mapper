using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using NPOI.SS.UserModel;
using Npoi.Mapper.Attributes;

namespace Npoi.Mapper
{
    /// <summary>
    /// Import Excel row data as object.
    /// </summary>
    public class Importer
    {
        private const BindingFlags BindingFlag = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance;

        #region Properties

        // Excel file workbook.
        public IWorkbook Workbook { get; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="Importer"/> class.
        /// </summary>
        /// <param name="stream">The input Excel(XLS, XLSX) file stream</param>
        public Importer(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            Workbook = WorkbookFactory.Create(stream, ImportOption.SheetContentOnly);
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Importer"/> class.
        /// </summary>
        /// <param name="workbook">The input IWorkbook object.</param>
        public Importer(IWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Workbook = workbook;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Get objects of target type by converting rows in the sheet with specified index (zero based).
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="sheetIndex">The sheet index; default is 0.</param>
        /// <param name="maxErrorRows">The maximum error rows before stop reading; default is 10.</param>
        /// <param name="objectInitializer">Factory method to create a new target object.</param>
        /// <returns>Objects of target type</returns>
        public IEnumerable<RowInfo<T>> TakeByHeader<T>(
            int sheetIndex = 0,
            int maxErrorRows = 10,
            Func<T> objectInitializer = null)
        {
            return TakeByHeader(Workbook.GetSheetAt(sheetIndex), maxErrorRows, objectInitializer);
        }

        /// <summary>
        /// Get objects of target type by converting rows in the sheet with specified name.
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="maxErrorRows">The maximum error rows before stopping read; default is 10.</param>
        /// <param name="objectInitializer">Factory method to create a new target object.</param>
        /// <returns>Objects of target type</returns>
        public IEnumerable<RowInfo<T>> TakeByHeader<T>(
            string sheetName,
            int maxErrorRows = 10,
            Func<T> objectInitializer = null)
        {
            return TakeByHeader(Workbook.GetSheet(sheetName), maxErrorRows, objectInitializer);
        }

        #endregion

        #region Private Methods

        private static IEnumerable<RowInfo<T>> TakeByHeader<T>(ISheet sheet, int maxErrorRows, Func<T> objectInitializer = null)
        {
            var list = new List<RowInfo<T>>();

            if (sheet == null || sheet.PhysicalNumberOfRows < 2)
            {
                return list;
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

                list.Add(data);

                if (data.ErrorColumnIndex >= 0) errorCount++;
            }

            return list;
        }

        private static void PrepareHeaders<T>(IRow headerRow, ICollection<ColumnInfo<T>> columns)
        {
            //
            // Column mapping priority:
            // ColumnNameAttribute > ColumnIndexAttribute > naming convention > MultiColumnsContainerAttribute
            //

            // Prepare a list of ColumnInfo.
            foreach (ICell header in headerRow)
            {
                // ColumnNameAttribute
                var column = GetColumnInfoByColumnsNameAttribute<T>(header);

                // ColumnIndexAttribute
                if (column == null)
                {
                    column = GetColumnInfoByColumnsIndexAttribute<T>(header);
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

                if (column != null)
                {
                    column.UseLastNonBlankValue = column.Property
                        .GetCustomAttributes<UseLastNonBlankValueAttribute>().Any();
                    columns.Add(column);
                }
            }
        }

        private static ColumnInfo<T> GetColumnInfoByName<T>(string name, int index)
        {
            var type = typeof(T);

            // First attempt: search by exact string.
            var pi = type.GetProperty(name, BindingFlag);
            if (pi != null) return new ColumnInfo<T>(name, index, pi);

            // Second attempt: search display name of DisplayAttribute if any.
            foreach (var propertyInfo in type.GetProperties(BindingFlag))
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

        private static ColumnInfo<T> GetColumnInfoByColumnsNameAttribute<T>(ICell header)
        {
            if (GetCellType(header) != CellType.String) return null;

            var type = typeof(T);

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var att = pi.GetCustomAttributes<ColumnNameAttribute>().FirstOrDefault();

                if (att == null) continue;

                if (string.Equals(att.Name, header.StringCellValue, StringComparison.CurrentCultureIgnoreCase))
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

        private static ColumnInfo<T> GetColumnInfoByColumnsIndexAttribute<T>(ICell header)
        {
            if (GetCellType(header) != CellType.String) return null;

            var type = typeof(T);

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var att = pi.GetCustomAttributes<ColumnIndexAttribute>().FirstOrDefault();

                if (att != null && att.Index == header.ColumnIndex)
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

        private static ColumnInfo<T> GetColumnInfoByMultiColumnsContainerAttribute<T>(ICell header)
        {
            var type = typeof(T);

            foreach (var pi in type.GetProperties(BindingFlag))
            {
                var att = pi.GetCustomAttributes<MultiColumnContainerAttribute>().FirstOrDefault();

                if (att == null) continue;

                var resolver = Activator.CreateInstance(att.ColumnResolverType) as ColumnResolver<T>;

                if (resolver == null) continue;

                var headerValue = GetHeaderValue(header);
                if (!resolver.TryResolveHeader(ref headerValue, header.ColumnIndex)) continue;

                return new ColumnInfo<T>(headerValue, header.ColumnIndex, pi, true)
                {
                    Resolver = resolver
                };
            }

            return null;
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
                    if (column.Property == null)
                    {
                        continue;
                    }

                    var cell = row.GetCell(column.Index);
                    var propertyType = column.Property.PropertyType;
                    object valueObj;

                    if (!TryGetCellValue(cell, propertyType, out valueObj))
                    {
                        errorIndex = column.Index;
                        errorMessage = "CellType is not supported yet!";
                        break;
                    }

                    valueObj = column.RefreshAndGetValue(valueObj);

                    if (column.Resolver != null)
                    {
                        if (!column.Resolver.TryResolveCell(column, valueObj, obj))
                        {
                            errorIndex = column.Index;
                            errorMessage = "Returned failure by custom cell resolver!";
                            break;
                        }
                    }
                    else if (valueObj != null)
                    {
                        // Change types between IConvertible objects, such as double, float, int and etc.
                        var value = Convert.ChangeType(valueObj, propertyType);
                        column.Property.SetValue(obj, value);
                    }
                    else
                    {
                        // If we go this far, keep target property untouched...
                    }
                }
                catch (Exception e)
                {
                    errorIndex = column.Index;
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
                    if (targetType.IsEnum) // Enum type.
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

        #endregion
    }
}
