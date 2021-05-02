using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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
    public class Mapper
    {
        #region Fields

        // Current working workbook.
        private IWorkbook _workbook;

        private Func<IColumnInfo, bool> _columnFilter;
        private Func<IColumnInfo, object, bool> _defaultTakeResolver;
        private Func<IColumnInfo, object, bool> _defaultPutResolver;
        private Action<ICell> _headerAction;

        #endregion

        #region Properties

        // Instance of helper class.
        private MapHelper Helper = new MapHelper();

        // Stores formats for type rather than specific property.
        internal readonly Dictionary<Type, string> TypeFormats = new Dictionary<Type, string>();

        // PropertyInfo map to ColumnAttribute
        private Dictionary<PropertyInfo, ColumnAttribute> Attributes { get; } = new Dictionary<PropertyInfo, ColumnAttribute>();

        // Property name map to ColumnAttribute
        private Dictionary<string, ColumnAttribute> DynamicAttributes { get; } = new Dictionary<string, ColumnAttribute>();

        /// <summary>
        /// Cache the tracked <see cref="ColumnInfo"/> objects by sheet name and target type.
        /// </summary>
        private Dictionary<string, Dictionary<Type, List<object>>> TrackedColumns { get; } =
            new Dictionary<string, Dictionary<Type, List<object>>>();

        /// <summary>
        /// Sheet name map to tracked objects in dictionary with row number as key.
        /// </summary>
        public Dictionary<string, Dictionary<int, object>> Objects { get; } = new Dictionary<string, Dictionary<int, object>>();

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
                    Helper.ClearCache();
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
        /// Gets or sets a zero-based index for the first row. It will be auto-detected if not set.
        /// If <see cref="HasHeader"/> is true (by default), this represents the header row index.
        /// </summary>
        public int FirstRowIndex { get; set; } = -1;

        /// <summary>
        /// Gets or sets a value indicating whether to skip blank rows when reading from Excel files. Default is true.
        /// </summary>
        /// <value>
        ///   <c>true</c> if blank lines are skipped; otherwise, <c>false</c>.
        /// </value>
        public bool SkipBlankRows { get; set; } = true;

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
        /// Use this to include and map columns for custom complex resolution.
        /// </summary>
        /// <param name="columnFilter">The function to determine whether or not to resolve an unmapped column.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Map(Func<IColumnInfo, bool> columnFilter, Func<IColumnInfo, object, bool> tryTake = null, Func<IColumnInfo, object, bool> tryPut = null)
        {
            _columnFilter = columnFilter;
            _defaultPutResolver = tryPut;
            _defaultTakeResolver = tryTake;

            return this;
        }

        /// <summary>
        /// Map by a <see cref="ColumnAttribute"/> object.
        /// </summary>
        /// <param name="attribute">The <see cref="ColumnAttribute"/> object.</param>
        /// <returns>The mapper object.</returns>
        public Mapper Map(ColumnAttribute attribute)
        {
            if (attribute == null) throw new ArgumentNullException(nameof(attribute));

            if (attribute.Property != null)
            {
                attribute.MergeTo(Attributes);
            }
            else if (attribute.PropertyName != null)
            {
                if (DynamicAttributes.ContainsKey(attribute.PropertyName))
                {
                    DynamicAttributes[attribute.PropertyName].MergeFrom(attribute);
                }
                else
                {
                    // Ensures column name for the first time mapping.
                    if (attribute.Name == null)
                    {
                        attribute.Name = attribute.PropertyName;
                    }

                    DynamicAttributes[attribute.PropertyName] = attribute;
                }
            }
            else
            {
                throw new InvalidOperationException("Either PropertyName or Property should be specified for a valid mapping!");
            }

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

        /* 
         * Removed this method in v3 since this is rarely used but just increased internal complexity.
         * 
        /// <summary>
        /// Specify the built-in format.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="builtinFormat">The built-in format, see https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html for possible values.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        [Obsolete("Builtin format will not be supported in next major release!")]
        public Mapper Format<T>(short builtinFormat, Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) return this;
            new ColumnAttribute { Property = pi, BuiltinFormat = builtinFormat }.MergeTo(Attributes);

            return this;
        }
        */

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
        /// Sets an action to configure header cells for export.
        /// </summary>
        /// <param name="headerAction">Action to configure header cell.</param>
        /// <returns>The mapper object.</returns>
        public Mapper ForHeader(Action<ICell> headerAction)
        {
            _headerAction = headerAction;
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
        public IEnumerable<RowInfo<T>> Take<T>(int sheetIndex = 0, int maxErrorRows = 10, Func<T> objectInitializer = null) where T : class
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
        public IEnumerable<RowInfo<T>> Take<T>(string sheetName, int maxErrorRows = 10, Func<T> objectInitializer = null) where T : class
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
            var sheet = Workbook.NumberOfSheets > sheetIndex ? Workbook.GetSheetAt(sheetIndex) : Workbook.CreateSheet();
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

        private IEnumerable<RowInfo<T>> Take<T>(ISheet sheet, int maxErrorRows, Func<T> objectInitializer = null) where T : class
        {
            if (sheet == null || sheet.PhysicalNumberOfRows < 1)
            {
                yield break;
            }

            var firstRowIndex = GetFirstRowIndex(sheet);
            var firstRow = sheet.GetRow(firstRowIndex);

            var targetType = typeof(T);
            if (targetType == typeof(object)) // Dynamic type.
            {
                targetType = GetDynamicType(sheet);
                MapHelper.LoadDynamicAttributes(Attributes, DynamicAttributes, targetType);
            }

            // Scan object attributes.
            MapHelper.LoadAttributes(Attributes, targetType);

            // Read the first row to get column information.
            var columns = GetColumns(firstRow, targetType);

            // Detect column format based on the first non-null cell.
            Helper.LoadDataFormats(sheet, HasHeader ? firstRowIndex + 1 : firstRowIndex, columns, TypeFormats);

            if (TrackObjects) Objects[sheet.SheetName] = new Dictionary<int, object>();

            // Loop rows in file. Generate one target object for each row.
            var errorCount = 0;
            var firstDataRowIndex = HasHeader ? firstRowIndex + 1 : firstRowIndex;
            foreach (IRow row in sheet)
            {
                if (maxErrorRows > 0 && errorCount >= maxErrorRows) break;
                if (row.RowNum < firstDataRowIndex) continue;

                if (SkipBlankRows && row.Cells.All(c => IsCellBlank(c))) continue;

                var obj = objectInitializer == null ? Activator.CreateInstance(targetType) : objectInitializer();
                var rowInfo = new RowInfo<T>(row.RowNum, obj as T, -1, string.Empty);
                LoadRowData(columns, row, obj, rowInfo);

                if (rowInfo.ErrorColumnIndex >= 0)
                {
                    errorCount++;
                    //rowInfo.Value = default(T);
                }
                if (TrackObjects) Objects[sheet.SheetName][row.RowNum] = rowInfo.Value;

                yield return rowInfo;
            }
        }

        private Type GetDynamicType(ISheet sheet)
        {
            var firstRowIndex = GetFirstRowIndex(sheet);
            var firstRow = sheet.GetRow(firstRowIndex);

            var names = new Dictionary<string, Type>();

            foreach (var header in firstRow)
            {
                var column = GetColumnInfoByDynamicAttribute(header);
                var type = Helper.InferColumnDataType(sheet, HasHeader ? firstRowIndex : -1, header.ColumnIndex);

                if (column != null)
                {
                    names[column.Attribute.PropertyName] = type ?? typeof(string);
                }
                else
                {
                    var headerValue = GetHeaderValue(header);
                    var tempColumn = new ColumnInfo(headerValue, header.ColumnIndex, null);
                    if (_columnFilter != null && !_columnFilter(tempColumn))
                    {
                        continue;
                    }

                    string propertyName;
                    if (HasHeader && MapHelper.GetCellType(header) == CellType.String)
                    {
                        propertyName = MapHelper.GetVariableName(header.StringCellValue, IgnoredNameChars,
                            TruncateNameFrom, header.ColumnIndex);
                    }
                    else
                    {
                        propertyName = MapHelper.GetVariableName(null, null, null, header.ColumnIndex);
                    }

                    names[propertyName] = type ?? typeof(string);
                    DynamicAttributes[propertyName] = new ColumnAttribute((ushort)header.ColumnIndex) { PropertyName = propertyName };
                }
            }

            return AnonymousTypeFactory.CreateType(names, true);
        }

        private static bool IsCellBlank(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String: return string.IsNullOrWhiteSpace(cell.StringCellValue);
                case CellType.Blank: return true;
                default: return false;
            };
        }

        private List<ColumnInfo> GetColumns(IRow headerRow, Type type)
        {
            //
            // Column mapping priority:
            // Map<T> > ColumnAttribute > naming convention > column filter.
            //

            var sheetName = headerRow.Sheet.SheetName;
            var columns = new List<ColumnInfo>();
            var columnsCache = new List<object>(); // Cached for export usage.

            // Prepare a list of ColumnInfo by the first row.
            foreach (ICell header in headerRow)
            {
                // Custom mappings via attributes.
                var column = GetColumnInfoByAttribute(header, type);

                // Naming convention.
                if (column == null && HasHeader && MapHelper.GetCellType(header) == CellType.String)
                {
                    var s = header.StringCellValue;

                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        column = GetColumnInfoByName(s.Trim(), header.ColumnIndex, type);
                    }
                }

                // Column filter.
                if (column == null)
                {
                    column = GetColumnInfoByFilter(header, _columnFilter);

                    if (column != null) // Set default resolvers since the column is not mapped explicitly.
                    {
                        column.Attribute.TryPut = _defaultPutResolver;
                        column.Attribute.TryTake = _defaultTakeResolver;
                    }
                }

                if (column == null) continue; // No property was mapped to this column.

                if (header.CellStyle != null) column.HeaderFormat = header.CellStyle.DataFormat;
                columns.Add(column);
                columnsCache.Add(column);
            }

            var typeDict = TrackedColumns.ContainsKey(sheetName)
                ? TrackedColumns[sheetName]
                : TrackedColumns[sheetName] = new Dictionary<Type, List<object>>();

            typeDict[type] = columnsCache;

            return columns;
        }

        private ColumnInfo GetColumnInfoByDynamicAttribute(ICell header)
        {
            var cellType = MapHelper.GetCellType(header);
            var index = header.ColumnIndex;

            foreach (var pair in DynamicAttributes)
            {
                var attribute = pair.Value;

                // If no header, cannot get a ColumnInfo by resolving header string.
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
                    return new ColumnInfo(headerValue, attribute);
                }
            }

            return null;
        }

        private ColumnInfo GetColumnInfoByAttribute(ICell header, Type type)
        {
            var cellType = MapHelper.GetCellType(header);
            var index = header.ColumnIndex;

            foreach (var pair in Attributes)
            {
                var attribute = pair.Value;

                if (pair.Key.ReflectedType != type || attribute.Ignored == true) continue;

                // If no header, cannot get a ColumnInfo by resolving header string.
                if (!HasHeader && attribute.Index < 0) continue;

                var headerValue = HasHeader ? GetHeaderValue(header) : null;
                var indexMatch = attribute.Index == index;
                var nameMatch = cellType == CellType.String &&
                                string.Equals(attribute.Name?.Trim(), header.StringCellValue?.Trim());

                // Index takes precedence over Name.
                if (indexMatch || (attribute.Index < 0 && nameMatch))
                {
                    // Use a clone so no pollution to original attribute,
                    // The origin might be used later again for multi-column/DefaultResolverType purpose.
                    attribute = attribute.Clone(index);
                    return new ColumnInfo(headerValue, attribute);
                }
            }

            return null;
        }

        private ColumnInfo GetColumnInfoByName(string name, int index, Type type)
        {
            // First attempt: search by string (ignore case).
            var pi = type.GetProperty(name, MapHelper.BindingFlag);

            if (pi == null)
            {
                // Second attempt: search display name of DisplayAttribute if any.
                foreach (var propertyInfo in type.GetProperties(MapHelper.BindingFlag))
                {
                    var attributes = propertyInfo.GetCustomAttributes(typeof(DisplayAttribute), false);

                    if (attributes.Any(att => string.Equals(((DisplayAttribute)att).Name, name, StringComparison.CurrentCultureIgnoreCase)))
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

            return attribute == null ? new ColumnInfo(name, index, pi) : new ColumnInfo(name, attribute);
        }

        private static ColumnInfo GetColumnInfoByFilter(ICell header, Func<IColumnInfo, bool> columnFilter)
        {
            if (columnFilter == null) return null;

            var headerValue = GetHeaderValue(header);
            var column = new ColumnInfo(headerValue, header.ColumnIndex, null);

            return !columnFilter(column) ? null : column;
        }

        private static void LoadRowData(IEnumerable<ColumnInfo> columns, IRow row, object target, IRowInfo rowInfo)
        {
            var errorIndex = -1;
            string errorMessage = null;

            void ColumnFailed(IColumnInfo column, string message)
            {
                if (errorIndex >= 0) return; // Ensures the first error will not be overwritten.
                if (column.Attribute.IgnoreErrors == true) return;
                errorIndex = column.Attribute.Index;
                errorMessage = message;
            }

            foreach (var column in columns)
            {
                var index = column.Attribute.Index;
                if (index < 0) continue;

                try
                {
                    var cell = row.GetCell(index);
                    var propertyType = column.Attribute.PropertyUnderlyingType ?? column.Attribute.Property?.PropertyType;

                    if (!MapHelper.TryGetCellValue(cell, propertyType, out object valueObj))
                    {
                        ColumnFailed(column, "CellType is not supported yet!");
                        continue;
                    }

                    valueObj = column.RefreshAndGetValue(valueObj);

                    if (column.Attribute.TryTake != null)
                    {
                        if (!column.Attribute.TryTake(column, target))
                        {
                            ColumnFailed(column, "Returned failure by custom cell resolver!");
                        }
                    }
                    else if (propertyType != null)
                    {
                        // Change types between IConvertible objects, such as double, float, int and etc.
                        if (MapHelper.TryConvertType(valueObj, column, out object result))
                        {
                            column.Attribute.Property.SetValue(target, result, null);
                        }
                        else
                        {
                            ColumnFailed(column, "Cannot convert value to the property type!");
                        }
                        //var value = Convert.ChangeType(valueObj, column.Attribute.PropertyUnderlyingType ?? propertyType);
                    }
                }
                catch (Exception e)
                {
                    ColumnFailed(column, e.ToString());
                }
            }

            rowInfo.ErrorColumnIndex = errorIndex;
            rowInfo.ErrorMessage = errorMessage;
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
            var firstRowIndex = GetFirstRowIndex(sheet);
            var firstRow = sheet.GetRow(firstRowIndex);
            var objectArray = objects as T[] ?? objects.ToArray();
            var type = MapHelper.GetConcreteType(objectArray);

            var columns = GetTrackedColumns(sheetName, type) ??
                           GetColumns(firstRow ?? PopulateFirstRow(sheet, null, type), type);
            firstRow = sheet.GetRow(firstRowIndex) ?? PopulateFirstRow(sheet, columns, type);

            var rowIndex = overwrite
                ? HasHeader ? firstRowIndex + 1 : firstRowIndex
                : sheet.GetRow(sheet.LastRowNum) != null ? sheet.LastRowNum + 1 : sheet.LastRowNum;

            MapHelper.EnsureDefaultFormats(columns, TypeFormats);

            foreach (var o in objectArray)
            {
                var row = sheet.GetRow(rowIndex);

                if (overwrite && row != null)
                {
                    //sheet.RemoveRow(row);
                    //row = sheet.CreateRow(rowIndex);
                    row.Cells.Clear();
                }

                row = row ?? sheet.CreateRow(rowIndex);

                foreach (var column in columns)
                {
                    var pi = column.Attribute.Property;
                    var value = pi?.GetValue(o, null);
                    var cell = row.GetCell(column.Attribute.Index, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    column.CurrentValue = value;
                    if (column.Attribute.TryPut == null || column.Attribute.TryPut(column, o))
                    {
                        SetCell(cell, column.CurrentValue, column, setStyle: overwrite);
                    }
                }

                rowIndex++;
            }

            // Remove not used rows if any.
            while (overwrite && rowIndex <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    //sheet.RemoveRow(row);
                    row.Cells.Clear();
                }
                rowIndex++;
            }

            // Injects custom action for headers.
            if (overwrite && HasHeader && _headerAction != null)
            {
                firstRow?.Cells.ForEach(c => _headerAction(c));
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

        private IRow PopulateFirstRow(ISheet sheet, List<ColumnInfo> columns, Type type)
        {
            var row = sheet.CreateRow(GetFirstRowIndex(sheet));

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

            MapHelper.LoadAttributes(Attributes, type);

            var attributes = Attributes.Where(p => p.Value.Property != null && p.Value.Property.ReflectedType == type);
            var properties = new List<PropertyInfo>(type.GetProperties(MapHelper.BindingFlag));

            // Firstly populate for those have Attribute specified.
            foreach (var pair in attributes)
            {
                var pi = pair.Key;
                var attribute = pair.Value;
                if (pair.Value.Index < 0) continue;

                var cell = row.CreateCell(attribute.Index);
                if (HasHeader)
                {
                    cell.SetCellValue(attribute.Name ?? pi.Name);
                }
                properties.Remove(pair.Key); // Remove populated property.
            }

            var index = 0;

            // Then populate for those do not have Attribute specified.
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

        private List<ColumnInfo> GetTrackedColumns(string sheetName, Type type)
        {
            if (!TrackedColumns.ContainsKey(sheetName)) return null;

            IEnumerable<ColumnInfo> columns = null;

            var cols = TrackedColumns[sheetName];
            if (cols.ContainsKey(type))
            {
                columns = cols[type].OfType<ColumnInfo>();
            }

            return columns?.ToList();
        }

        private void SetCell(ICell cell, object value, ColumnInfo column, bool isHeader = false, bool setStyle = true)
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

            if (column != null && setStyle)
            {
                column.SetCellStyle(cell, value, isHeader, TypeFormats, Helper);
            }
        }

        #endregion Export

        private int GetFirstRowIndex(ISheet sheet) => FirstRowIndex >= 0 ? FirstRowIndex : sheet.FirstRowNum;

        #endregion
    }
}
