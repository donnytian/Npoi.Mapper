using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Npoi.Mapper.Attributes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Npoi.Mapper
{
    /// <summary>
    /// Map object properties with Excel columns for importing from and exporting to file.
    /// </summary>
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    [SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Global")]
    public class Mapper
    {
        #region Fields

        // Current working workbook.
        private IWorkbook _workbook;

        #endregion

        #region Properties

        // PropertyInfo map to ColumnAttribute
        private Dictionary<PropertyInfo, ColumnAttribute> Attributes { get; } = new Dictionary<PropertyInfo, ColumnAttribute>();

        /// <summary>
        /// Cache the tracked <see cref="ColumnInfo{TTarget}"/> objects by sheet name and target type.
        /// </summary>
        private Dictionary<string, Dictionary<Type, List<object>>> TrackedColumns { get; } =
            new Dictionary<string, Dictionary<Type, List<object>>>();

        /// <summary>
        /// Sheet name map to tracked objects in dictionary with row number as key.
        /// </summary>
        public Dictionary<string, Dictionary<int, object>> Objects { get; } = new Dictionary<string, Dictionary<int, object>>();

        /// <summary>
        /// Type of class that implemented <see cref="IColumnResolver{TTarget}"/> to handle unrecognized/ unmapped columns.
        /// </summary>
        public Type DefaultResolverType { get; set; }

        /// <summary>
        /// The Excel workbook.
        /// </summary>
        public IWorkbook Workbook
        {
            get { return _workbook; }

            private set
            {
                if (value != _workbook)
                {
                    Objects.Clear();
                    TrackedColumns.Clear();
                    MapHelper.ClearCache();
                }
                _workbook = value;
            }
        }

        /// <summary>
        /// When map column, chars in this array will be removed from column header.
        /// </summary>
        public char[] IgnoredNameChars { get; set; }

        /// <summary>
        /// When map column, column name will be truncated from any of chars in this array.
        /// </summary>
        public char[] TruncateNameFrom { get; set; }

        /// <summary>
        /// Whether to track objects read from the Excel file. Default is true.
        /// If object tracking is enabled, the <see cref="Mapper"/> object keeps track of objects it yields through the Take() methods.
        /// You can then modify these objects and save them back to an Excel file without having to specify the list of objects to save.
        /// </summary>
        public bool TrackObjects { get; set; } = true;

        /// <summary>
        /// Whether to take the first row as column header. Default is true.
        /// </summary>
        public bool HasHeader { get; set; } = true;

        /// <summary>
        /// Set a zero-based row index for header. It will be auto-detected if not set.
        /// </summary>
        public int HeaderRowIndex { get; set; } = -1; // TODO: enable this...

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class.
        /// </summary>
        public Mapper()
        {
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class with stream to read workbook.
        /// </summary>
        /// <param name="stream">The input Excel(XLS, XLSX) file stream</param>
        public Mapper(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            using (stream)
            {
                Workbook = WorkbookFactory.Create(stream);
            }
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class with a workbook.
        /// </summary>
        /// <param name="workbook">The input IWorkbook object.</param>
        public Mapper(IWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Workbook = workbook;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="Mapper"/> class with file path to read workbook.
        /// </summary>
        /// <param name="filePath">The path of Excel file.</param>
        public Mapper(string filePath) : this(new FileStream(filePath, FileMode.Open))
        {
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Map property to a column by specified column name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public Mapper Map<T>(string columnName, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            if (columnName == null)
                throw new ArgumentNullException(nameof(columnName));

            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            new ColumnAttribute
            {
                Property = pi,
                Name = columnName,
                ResolverType = resolverType,
                Ignored = false
            }.MergeTo(Attributes);

            return this;
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based).
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public Mapper Map<T>(ushort columnIndex, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;

            new ColumnAttribute
            {
                Property = pi,
                Index = columnIndex,
                ResolverType = resolverType,
                Ignored = false
            }.MergeTo(Attributes);

            return this;
        }

        /// <summary>
        /// Specify to use last non-blank value from above cell for a property.
        /// Typically to address the blank cell issue in merged cells.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper UseLastNonBlankValue<T>(Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;
            new ColumnAttribute { Property = pi, UseLastNonBlankValue = true }.MergeTo(Attributes);

            return this;
        }

        /// <summary>
        /// Specify to ignore a property. Ignored property will not be mapped for import and export.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Ignore<T>(Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;
            new ColumnAttribute { Property = pi, Ignored = true }.MergeTo(Attributes);

            return this;
        }

        /// <summary>
        /// Specify the built-in format.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="builtinFormat">The built-in format, see https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html for possible values.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Format<T>(short builtinFormat, Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;
            new ColumnAttribute { Property = pi, BuiltinFormat = builtinFormat }.MergeTo(Attributes);

            return this;
        }

        /// <summary>
        /// Specify the custom format.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="customFormat">The custom format, see https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 for the syntax.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Format<T>(string customFormat, Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;
            new ColumnAttribute { Property = pi, CustomFormat = customFormat }.MergeTo(Attributes);

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
            return Take(sheet, maxErrorRows, objectInitializer);
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
            return Take(sheet, maxErrorRows, objectInitializer);
        }

        /// <summary>
        /// Put objects in the sheet with specified name.
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing rows; otherwise append.</param>
        public void Put<T>(IEnumerable<T> objects, string sheetName, bool overwrite = true)
        {
            if (Workbook == null) Workbook = new XSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName) ?? Workbook.CreateSheet(sheetName);
            Put(sheet, objects, overwrite);
        }

        /// <summary>
        /// Put objects in the sheet with specified zero-based index.
        /// </summary>
        /// <typeparam name="T">Target object type</typeparam>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">The sheet index, default is 0.</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing rows; otherwise append.</param>
        public void Put<T>(IEnumerable<T> objects, int sheetIndex = 0, bool overwrite = true)
        {
            if (Workbook == null) Workbook = new XSSFWorkbook();
            var sheet = Workbook.GetSheetAt(sheetIndex);
            Put(sheet, objects, overwrite);
        }

        /// <summary>
        /// Saves the current workbook to specified file.
        /// </summary>
        /// <param name="path">The path to the Excel file.</param>
        public void Save(string path)
        {
            if (Workbook == null) return;

            using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
                Workbook.Write(fs);
        }

        /// <summary>
        /// Saves the current workbook to specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the workbook.</param>
        public void Save(Stream stream)
        {
            Workbook?.Write(stream);
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="path">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="overwrite">If file exists, pass <c>true</c> to overwrite existing file; <c>false</c> to merge.</param>
        /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(string path, IEnumerable<T> objects, string sheetName, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null && !overwrite) LoadWorkbookFromFile(path);

            using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
                Save(fs, objects, sheetName, overwrite, xlsx);
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="path">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="overwrite">If file exists, pass <c>true</c> to overwrite existing file; <c>false</c> to merge.</param>
        /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(string path, IEnumerable<T> objects, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null && !overwrite) LoadWorkbookFromFile(path);

            using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
                Save(fs, objects, sheetIndex, overwrite, xlsx);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to write the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName) ?? Workbook.CreateSheet(sheetName);
            Save(stream, sheet, objects, overwrite);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to write the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.NumberOfSheets > sheetIndex ? Workbook.GetSheetAt(sheetIndex) : Workbook.CreateSheet();
            Save(stream, sheet, objects, overwrite);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="path">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="overwrite">If file exists, pass <c>true</c> to overwrite existing file; <c>false</c> to merge.</param>
        /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(string path, string sheetName, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null && !overwrite) LoadWorkbookFromFile(path);

            using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
                Save<T>(fs, sheetName, overwrite, xlsx);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="path">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="overwrite">If file exists, pass <c>true</c> to overwrite existing file; <c>false</c> to merge.</param>
        public void Save<T>(string path, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null && !overwrite) LoadWorkbookFromFile(path);

            using (var fs = File.Open(path, FileMode.Create, FileAccess.Write))
                Save<T>(fs, sheetIndex, overwrite, xlsx);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to write the objects to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
        /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, string sheetName, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName) ?? Workbook.CreateSheet(sheetName);

            Save<T>(stream, sheet, overwrite);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to write the objects to.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheetAt(sheetIndex) ?? Workbook.CreateSheet();

            Save<T>(stream, sheet, overwrite);
        }

        #endregion

        #region Private Methods

        #region Import

        private IEnumerable<RowInfo<T>> Take<T>(ISheet sheet, int maxErrorRows, Func<T> objectInitializer = null)
        {
            if (sheet == null || sheet.PhysicalNumberOfRows < 1)
            {
                yield break;
            }

            var firstRowIndex = sheet.FirstRowNum;
            var firstRow = sheet.GetRow(firstRowIndex);

            // Scan object attributes.
            MapHelper.LoadAttributes<T>(Attributes);

            // Read the first row to get column information.
            var columns = GetColumns<T>(firstRow);
            MapHelper.LoadDataFormats(sheet.GetRow(firstRowIndex + 1), columns);

            if (TrackObjects) Objects[sheet.SheetName] = new Dictionary<int, object>();

            // Loop rows in file. Generate one target object for each row.
            var errorCount = 0;
            foreach (IRow row in sheet)
            {
                if (maxErrorRows > 0 && errorCount >= maxErrorRows) break;
                if (HasHeader && row.RowNum == firstRowIndex) continue;

                var data = GetRowData(columns, row, objectInitializer);

                if (data.ErrorColumnIndex >= 0) errorCount++;
                if (TrackObjects) Objects[sheet.SheetName][row.RowNum] = data.Value;

                yield return data;
            }
        }

        private List<ColumnInfo<T>> GetColumns<T>(IRow headerRow)
        {
            //
            // Column mapping priority:
            // Map<T> > ColumnAttribute > naming convention > DefaultResolverType.
            //

            var sheetName = headerRow.Sheet.SheetName;
            var columns = new List<ColumnInfo<T>>();
            var columnsCache = new List<object>(); // Cached for export usage.

            // Prepare a list of ColumnInfo by the first row.
            foreach (ICell header in headerRow)
            {
                // Custom mappings via attributes.
                var column = GetColumnInfoByAttribute<T>(header);

                // Naming convention.
                if (column == null && HasHeader && MapHelper.GetCellType(header) == CellType.String)
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

                if (column == null) continue; // No property was mapped to this column.

                if (header.CellStyle != null) column.HeaderFormat = header.CellStyle.DataFormat;
                columns.Add(column);
                columnsCache.Add(column);
            }

            var typeDict = TrackedColumns.ContainsKey(sheetName)
                ? TrackedColumns[sheetName]
                : TrackedColumns[sheetName] = new Dictionary<Type, List<object>>();

            typeDict[typeof(T)] = columnsCache;

            return columns;
        }

        private ColumnInfo<T> GetColumnInfoByAttribute<T>(ICell header)
        {
            var type = typeof(T);
            var cellType = MapHelper.GetCellType(header);
            var index = header.ColumnIndex;

            foreach (var pair in Attributes)
            {
                var attribute = pair.Value;

                if (pair.Key.ReflectedType != type || attribute.Ignored == true) continue;

                // If no header, cannot get a ColumnInfo by resolving header via custom resolver.
                if (!HasHeader && attribute.Index < 0) continue;

                var headerValue = HasHeader ? GetHeaderValue(header) : null;
                var indexMatch = attribute.Index == index;
                var nameMatch = cellType == CellType.String && string.Equals(attribute.Name, header.StringCellValue);

                // Index takes precedence over Name.
                if (indexMatch || (attribute.Index < 0 && nameMatch))
                {
                    // Use a clone so no pollution to original attribute,
                    // The origin might be used later again for multi-column/DefaultResolverType purpose.
                    attribute = attribute.Clone(index);

                    var resolver = attribute.ResolverType == null ?
                        null :
                        Activator.CreateInstance(attribute.ResolverType) as IColumnResolver<T>;
                    resolver?.IsColumnMapped(ref headerValue, index); // Ignore return value since it's already mapped to column.

                    return new ColumnInfo<T>(headerValue, attribute)
                    {
                        Resolver = resolver
                    };
                }

                // If goes this far, try map column by custom resolver.
                if (attribute.Index < 0 && attribute.Name == null && attribute.ResolverType != null)
                {
                    var resolver = Activator.CreateInstance(attribute.ResolverType) as IColumnResolver<T>;

                    if (resolver == null) continue;

                    // Check if this column is desired by resolver.
                    if (!resolver.IsColumnMapped(ref headerValue, index)) continue;

                    attribute = attribute.Clone(index);

                    return new ColumnInfo<T>(headerValue, attribute)
                    {
                        Resolver = resolver
                    };
                }
            }

            return null;
        }

        private ColumnInfo<T> GetColumnInfoByName<T>(string name, int index)
        {
            var type = typeof(T);

            // First attempt: search by string (ignore case).
            var pi = type.GetProperty(name, MapHelper.BindingFlag);

            if (pi == null)
            {
                // Second attempt: search display name of DisplayAttribute if any.
                foreach (var propertyInfo in type.GetProperties(MapHelper.BindingFlag))
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
                // Third attempt: remove ignored chars and do the truncation.
                pi = type.GetProperty(MapHelper.GetRefinedName(name, IgnoredNameChars, TruncateNameFrom), MapHelper.BindingFlag);
            }

            if (pi == null) return null;

            ColumnAttribute attribute = null;

            if (Attributes.ContainsKey(pi))
            {
                attribute = Attributes[pi].Clone(index);
                if (attribute.Ignored == true) return null;
            }

            return attribute == null ? new ColumnInfo<T>(name, index, pi) : new ColumnInfo<T>(name, attribute);
        }

        private static ColumnInfo<T> GetColumnInfoByResolverType<T>(ICell header, Type resolverType)
        {
            if (resolverType == null) return null;

            var resolver = Activator.CreateInstance(resolverType) as IColumnResolver<T>;

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
                var index = column.Attribute.Index;
                if (index < 0) continue;

                try
                {
                    var cell = row.GetCell(index);
                    var propertyType = column.Attribute.Property?.PropertyType;
                    object valueObj;

                    if (!MapHelper.TryGetCellValue(cell, propertyType, out valueObj))
                    {
                        errorIndex = index;
                        errorMessage = "CellType is not supported yet!";
                        break;
                    }

                    valueObj = column.RefreshAndGetValue(valueObj);

                    if (column.Resolver != null)
                    {
                        if (!column.Resolver.TryTakeCell(column, valueObj, obj))
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

        private void LoadWorkbookFromFile(string path)
        {
            Workbook = WorkbookFactory.Create(new FileStream(path, FileMode.Open));
        }

        #endregion

        #region Export

        private void Put<T>(ISheet sheet, IEnumerable<T> objects, bool overwrite)
        {
            var sheetName = sheet.SheetName;
            var firstRow = sheet.GetRow(sheet.FirstRowNum);

            List<ColumnInfo<T>> columns = null;
            if (!overwrite) columns = GetTrackedColumns<T>(sheetName);
            if (columns == null) columns = GetColumns<T>(firstRow ?? PopulateFirstRow<T>(sheet));
            if (firstRow == null) PopulateFirstRow(sheet, columns);

            var rowIndex = overwrite
                ? HasHeader ? sheet.FirstRowNum + 1 : sheet.FirstRowNum
                : sheet.GetRow(sheet.LastRowNum) != null ? sheet.LastRowNum + 1 : sheet.LastRowNum;

            foreach (var o in objects)
            {
                var row = sheet.GetRow(rowIndex);

                if (overwrite && row != null)
                {
                    sheet.RemoveRow(row);
                    row = sheet.CreateRow(rowIndex);
                }

                row = row ?? sheet.CreateRow(rowIndex);

                foreach (var column in columns)
                {
                    var pi = column.Attribute.Property;
                    var value = pi?.GetValue(o);
                    var cell = row.GetCell(column.Attribute.Index, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    if (column.Resolver != null)
                    {
                        column.Resolver.TryPutCell(column, out value, o);
                    }
                    else
                    {
                        SetCell(cell, value, column, setStyle: overwrite);
                    }
                }

                rowIndex++;
            }

            // Remove not used rows if any.
            while (overwrite && rowIndex <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(rowIndex);
                if (row != null) sheet.RemoveRow(row);
                rowIndex++;
            }
        }

        private void Save<T>(Stream stream, ISheet sheet, bool overwrite)
        {
            var sheetName = sheet.SheetName;
            var objects = Objects.ContainsKey(sheetName) ? Objects[sheetName] : new Dictionary<int, object>();

            Put(sheet, objects.Values.OfType<T>(), overwrite);
            Workbook.Write(stream);
        }

        private void Save<T>(Stream stream, ISheet sheet, IEnumerable<T> objects, bool overwrite)
        {
            Put(sheet, objects, overwrite);
            Workbook.Write(stream);
        }

        private IRow PopulateFirstRow<T>(ISheet sheet, List<ColumnInfo<T>> columns = null)
        {
            var row = sheet.CreateRow(sheet.FirstRowNum);

            // Use existing column populate the first row.

            if (columns != null)
            {
                foreach (var column in columns)
                {
                    var cell = row.CreateCell(column.Attribute.Index);

                    if (!HasHeader) continue;

                    SetCell(cell, column.Attribute.Name ?? column.HeaderValue, column, true);
                }

                return row;
            }

            // If no column cached, populate the first row with attributes and object properties.

            var type = typeof(T);

            MapHelper.LoadAttributes<T>(Attributes);

            var attributes = Attributes.Where(p => p.Value.Property != null && p.Value.Property.ReflectedType == type);
            var properties = new List<PropertyInfo>(type.GetProperties(MapHelper.BindingFlag));

            // Firstly populate for those have attribute specified.
            foreach (var pair in attributes)
            {
                var pi = pair.Key;
                var attribute = pair.Value;
                if (pair.Value.Index < 0) continue;

                var cell = row.CreateCell(attribute.Index);
                if (HasHeader) cell.SetCellValue(attribute.Name ?? pi.Name);
                properties.Remove(pair.Key); // Remove populated property.
            }

            var index = 0;

            // Then populate for those do not have attribute specified.
            foreach (var pi in properties)
            {
                var attribute = Attributes.ContainsKey(pi) ? Attributes[pi] : null;
                if (attribute?.Ignored == true) continue;

                while (row.GetCell(index) != null) index++;
                var cell = row.CreateCell(index);
                if (HasHeader)
                {
                    cell.SetCellValue(attribute?.Name ?? pi.Name);
                }
                else
                {
                    new ColumnAttribute { Index = index, Property = pi }.MergeTo(Attributes);
                }
                index++;
            }

            return row;
        }

        private List<ColumnInfo<T>> GetTrackedColumns<T>(string sheetName)
        {
            if (!TrackedColumns.ContainsKey(sheetName)) return null;

            IEnumerable<ColumnInfo<T>> columns = null;

            var cols = TrackedColumns[sheetName];
            var type = typeof(T);
            if (cols.ContainsKey(type))
            {
                columns = cols[type].OfType<ColumnInfo<T>>();
            }

            return columns?.ToList();
        }

        private static void SetCell<T>(
            ICell cell,
            object value,
            ColumnInfo<T> column,
            bool isHeader = false,
            bool setStyle = true)
        {
            if (value == null || value is ICollection)
            {
                cell.SetCellValue((string)null);
            }
            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
            }
            else if (value.GetType().IsNumeric())
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else if (value is bool)
            {
                cell.SetCellValue((bool)value);
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }

            if (setStyle) column.SetCellStyle(cell, isHeader);
        }

        #endregion Export

        #endregion
    }
}
