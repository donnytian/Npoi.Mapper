using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Npoi.Mapper.Attributes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
// ReSharper disable MemberCanBePrivate.Global

namespace Npoi.Mapper;

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
    private readonly MapHelper _helper = new MapHelper();

    // Stores formats for type rather than specific property.
    internal readonly Dictionary<Type, string> TypeFormats = new();

    // PropertyInfo map to ColumnAttribute
    private Dictionary<string, ColumnAttribute> Attributes { get; } = new();

    // Property name map to ColumnAttribute
    private Dictionary<string, ColumnAttribute> DynamicAttributes { get; } = new();

    /// <summary>
    /// Cache the tracked <see cref="ColumnInfo"/> objects by sheet name and target type.
    /// </summary>
    private Dictionary<string, Dictionary<Type, List<object>>> TrackedColumns { get; } = new();

    /// <summary>
    /// The Excel workbook.
    /// </summary>
    public IWorkbook Workbook
    {
        get => _workbook;

        private set
        {
            if (value != _workbook)
            {
                TrackedColumns.Clear();
                _helper.ClearCache();

                if (value is HSSFWorkbook)
                {
                    FormulaEvaluator = new HSSFFormulaEvaluator(value);
                }
                else if (value is XSSFWorkbook)
                {
                    FormulaEvaluator = new XSSFFormulaEvaluator(value);
                }
            }
            _workbook = value;
        }
    }

    public IFormulaEvaluator FormulaEvaluator { get; private set; }

    /// <summary>
    /// When map column, chars in this array will be removed from column header.
    /// </summary>
    public char[] IgnoredNameChars { get; set; }

    /// <summary>
    /// When map column, column name will be truncated from any of chars in this array.
    /// </summary>
    public char[] TruncateNameFrom { get; set; }

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
    public bool SkipBlankRows { get; set; } = false;

    /// <summary>
    /// Gets or sets a value indicating whether to trim blanks from values in rows. Default is None.
    /// </summary>
    /// <value>
    ///   <c>Start</c> to trim initial spaces; <c>End</c> to trim end spaces; <c>Both</c> to trim initial and end spaces; <c>None</c> to preserve spaces in values.
    /// </value>
    public TrimSpacesType TrimSpaces { get; set; } = TrimSpacesType.None;

    /// <summary>
    /// Gets or sets a value indicating whether to read <see cref="DefaultValueAttribute"/> value and assume it as default value when excel column is blank. Default is false.
    /// </summary>
    /// <value>
    ///   <c>true</c> if <see cref="DefaultValueAttribute"/> is to be considered; otherwise, <c>false</c>.
    /// </value>
    public bool UseDefaultValueAttribute { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether to write default property values to excel. Default is false.
    /// </summary>
    /// <value>
    ///   <c>true</c> if default values result in empty cells in excel; otherwise, <c>false</c> means all values are written to excel, even if equal to default.
    /// </value>
    public bool SkipWriteDefaultValue { get; set; } = false;

    /// <summary>
    /// Gets or sets a value indicating whether to skip hidden rows when reading from Excel files. Default is false.
    /// </summary>
    /// <value>
    ///   <c>true</c> if hidden lines are skipped; otherwise, <c>false</c>.
    /// </value>
    public bool SkipHiddenRows { get; set; }

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
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
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
        if (attribute == null)
        {
            throw new ArgumentNullException(nameof(attribute));
        }

        if (attribute.PropertyType != null)
        {
            attribute.MergeTo(Attributes);
        }
        else if (attribute.PropertyName != null) // For dynamic type
        {
            if (DynamicAttributes.TryGetValue(attribute.PropertyName, out var dynamicAttribute))
            {
                dynamicAttribute.MergeFrom(attribute);
            }
            else
            {
                // Ensures column name for the first time mapping.
                attribute.Name ??= attribute.PropertyName;
                DynamicAttributes[attribute.PropertyName] = attribute;
            }
        }
        else
        {
            throw new InvalidOperationException("Either PropertyName or property selector should be specified for a valid mapping!");
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
        var (pi, fullPath) = MapHelper.GetPropertyInfo(propertySelector);
        new ColumnAttribute { UseLastNonBlankValue = true }.SetProperty(pi, fullPath).MergeTo(Attributes);

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
        var (pi, fullPath) = MapHelper.GetPropertyInfo(propertySelector);
        new ColumnAttribute { Ignored = true }.SetProperty(pi, fullPath).MergeTo(Attributes);

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
        var (pi, fullPath) = MapHelper.GetPropertyInfo(propertySelector);
        new ColumnAttribute { CustomFormat = customFormat }.SetProperty(pi, fullPath).MergeTo(Attributes);

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
    /// <param name="maxErrorRows">The maximum error rows before stop reading; default is 100.</param>
    /// <param name="objectInitializer">Factory method to create a new target object.</param>
    /// <returns>Objects of target type</returns>
    public IEnumerable<RowInfo<T>> Take<T>(int sheetIndex = 0, int maxErrorRows = 100, Func<T> objectInitializer = null) where T : class
    {
        var sheet = Workbook.GetSheetAt(sheetIndex);
        return Take(null, sheet, maxErrorRows, objectInitializer);
    }

    /// <summary>
    /// Get objects of target type by converting rows in the sheet with specified name.
    /// </summary>
    /// <typeparam name="T">Target object type</typeparam>
    /// <param name="sheetName">The sheet name</param>
    /// <param name="maxErrorRows">The maximum error rows before stopping read; default is 100.</param>
    /// <param name="objectInitializer">Factory method to create a new target object.</param>
    /// <returns>Objects of target type</returns>
    public IEnumerable<RowInfo<T>> Take<T>(string sheetName, int maxErrorRows = 100, Func<T> objectInitializer = null) where T : class
    {
        var sheet = Workbook.GetSheet(sheetName);
        return Take(null, sheet, maxErrorRows, objectInitializer);
    }

    /// <summary>
    /// Get objects as dynamic type with custom column type resolver.
    /// </summary>
    /// <param name="getColumnType">Function to get column type by inspecting the current column header.</param>
    /// <param name="sheetName">The sheet name.</param>
    /// <param name="maxErrorRows">The maximum error rows before stopping read; default is 100.</param>
    /// <returns>Objects of dynamic type.</returns>
    public IEnumerable<RowInfo<dynamic>> TakeDynamicWithColumnType(Func<ICell, Type> getColumnType, string sheetName, int maxErrorRows = 100)
    {
        var sheet = Workbook.GetSheet(sheetName);
        return Take<object>(getColumnType, sheet, maxErrorRows);
    }

    /// <summary>
    /// Get objects as dynamic type with custom column type resolver.
    /// </summary>
    /// <param name="getColumnType">Function to get column type by inspecting the current column header.</param>
    /// <param name="sheetIndex">The sheet index; default is 0.</param>
    /// <param name="maxErrorRows">The maximum error rows before stopping read; default is 100.</param>
    /// <returns>Objects of dynamic type.</returns>
    public IEnumerable<RowInfo<dynamic>> TakeDynamicWithColumnType(Func<ICell, Type> getColumnType, int sheetIndex = 0, int maxErrorRows = 100)
    {
        var sheet = Workbook.GetSheetAt(sheetIndex);
        return Take<object>(getColumnType, sheet, maxErrorRows);
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
        Workbook ??= new XSSFWorkbook();
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
        Workbook ??= new XSSFWorkbook();
        var sheet = Workbook.NumberOfSheets > sheetIndex ? Workbook.GetSheetAt(sheetIndex) : Workbook.CreateSheet();
        Put(sheet, objects, overwrite);
    }

    /// <summary>
    /// Saves the current workbook to specified file.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    public void Save(string path, bool leaveOpen)
    {
        if (Workbook == null) return;

        using var fileStream = GetStreamForSave(path);
        Workbook.Write(fileStream, leaveOpen);
    }

    /// <summary>
    /// Saves the current workbook to specified stream.
    /// </summary>
    /// <param name="stream">The stream to save the workbook.</param>
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    public void Save(Stream stream, bool leaveOpen)
    {
        Workbook?.Write(stream, leaveOpen);
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
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    public void Save<T>(string path, IEnumerable<T> objects, string sheetName, bool leaveOpen, bool overwrite = true, bool xlsx = true)
    {
        if (Workbook == null && File.Exists(path)) LoadWorkbookFromFile(path);

        using var fileStream = GetStreamForSave(path);
        Save(fileStream, objects, sheetName, leaveOpen, overwrite, xlsx);
    }

    /// <summary>
    /// Saves the specified objects to the specified Excel file.
    /// </summary>
    /// <typeparam name="T">The type of objects to save.</typeparam>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="objects">The objects to save.</param>
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    /// <param name="sheetIndex">Index of the sheet.</param>
    /// <param name="overwrite">If file exists, pass <c>true</c> to overwrite existing file; <c>false</c> to merge.</param>
    /// <param name="xlsx">if <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
    public void Save<T>(string path, IEnumerable<T> objects, bool leaveOpen, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
    {
        if (Workbook == null && File.Exists(path)) LoadWorkbookFromFile(path);

        using var fileStream = GetStreamForSave(path);
        Save(fileStream, objects, leaveOpen, sheetIndex, overwrite, xlsx);
    }

    /// <summary>
    /// Saves the specified objects to the specified stream.
    /// </summary>
    /// <typeparam name="T">The type of objects to save.</typeparam>
    /// <param name="stream">The stream to write the objects to.</param>
    /// <param name="objects">The objects to save.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
    /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
    public void Save<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool leaveOpen, bool overwrite = true, bool xlsx = true)
    {
        Workbook ??= xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
        var sheet = Workbook.GetSheet(sheetName) ?? Workbook.CreateSheet(sheetName);
        Save(stream, sheet, objects, leaveOpen, overwrite);
    }

    /// <summary>
    /// Saves the specified objects to the specified stream.
    /// </summary>
    /// <typeparam name="T">The type of objects to save.</typeparam>
    /// <param name="stream">The stream to write the objects to.</param>
    /// <param name="objects">The objects to save.</param>
    /// <param name="leaveOpen">True to leave the stream open after write.</param>
    /// <param name="sheetIndex">Index of the sheet.</param>
    /// <param name="overwrite"><c>true</c> to overwrite existing content; <c>false</c> to merge.</param>
    /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
    public void Save<T>(Stream stream, IEnumerable<T> objects, bool leaveOpen, int sheetIndex = 0, bool overwrite = true, bool xlsx = true)
    {
        Workbook ??= xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
        var sheet = Workbook.NumberOfSheets > sheetIndex ? Workbook.GetSheetAt(sheetIndex) : Workbook.CreateSheet();
        Save(stream, sheet, objects, leaveOpen, overwrite);
    }

    #endregion

    #region Private Methods

    #region Import

    private IEnumerable<RowInfo<T>> Take<T>(Func<ICell, Type> getColumnType, ISheet sheet, int maxErrorRows, Func<T> objectInitializer = null)
        where T : class
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
            targetType = GetDynamicType(sheet, getColumnType);
            MapHelper.LoadDynamicAttributes(Attributes, DynamicAttributes, targetType);
            DynamicAttributes.Clear(); // Avoid mixed with other sheet.
        }

        // Scan object attributes.
        MapHelper.LoadAttributes(Attributes, targetType);

        // Read the first row to get column information.
        var columns = GetColumns(firstRow, targetType);

        // Detect column format based on the first non-null cell.
        _helper.LoadDataFormats(sheet, HasHeader ? firstRowIndex + 1 : firstRowIndex, columns, TypeFormats);

        // Loop rows in file. Generate one target object for each row.
        var errorCount = 0;
        var firstDataRowIndex = HasHeader ? firstRowIndex + 1 : firstRowIndex;
        foreach (IRow row in sheet)
        {
            if (maxErrorRows > 0 && errorCount >= maxErrorRows) break;
            if (row.RowNum < firstDataRowIndex) continue;

            if (SkipHiddenRows && row.Hidden.HasValue && row.Hidden.Value) continue;
            if (SkipBlankRows && row.Cells.All(IsCellBlank)) continue;

            var obj = objectInitializer == null ? Activator.CreateInstance(targetType) : objectInitializer();
            var rowInfo = new RowInfo<T>(row.RowNum, obj as T, -1, string.Empty);
            LoadRowData(columns, row, obj as T, rowInfo);

            if (rowInfo.ErrorColumnIndex >= 0)
            {
                errorCount++;
                //rowInfo.Value = default(T);
            }

            yield return rowInfo;
        }
    }

    private Type GetDynamicType(ISheet sheet, Func<ICell, Type> getColumnType)
    {
        var firstRowIndex = GetFirstRowIndex(sheet);
        var firstRow = sheet.GetRow(firstRowIndex);

        var names = new Dictionary<string, Type>();

        foreach (var header in firstRow)
        {
            var column = GetColumnInfoByDynamicAttribute(header);
            var type = getColumnType?.Invoke(header)
                ?? _helper.InferColumnDataType(sheet, HasHeader ? firstRowIndex : -1, header.ColumnIndex);

            if (column != null)
            {
                names[column.Attribute.PropertyName] = type ?? typeof(string);
            }
            else
            {
                var headerValue = GetHeaderValue(header);
                var tempColumn = new ColumnInfo(headerValue, header.ColumnIndex);
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
        }
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

            if (column == null)
            {
                continue; // No property was mapped to this column.
            }

            if (header.CellStyle != null) column.HeaderFormat = header.CellStyle.DataFormat;
            columns.Add(column);
            columnsCache.Add(column);
        }

        var typeDict = TrackedColumns.TryGetValue(sheetName, out var trackedColumn)
            ? trackedColumn
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

            if (!pair.Key.StartsWith(type.Name + '.') || attribute.Ignored == true)
            {
                continue;
            }

            // If no header, cannot get a ColumnInfo by resolving header string.
            if (!HasHeader && attribute.Index < 0)
            {
                continue;
            }

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

        if (pi == null)
        {
            return null;
        }

        ColumnAttribute attribute = null;
        var key = type.Name + "." + pi.Name;

        if (Attributes.TryGetValue(key, out var attribute1))
        {
            if (attribute1.Ignored == true)
            {
                return null;
            }
            attribute = attribute1.Clone(index);
        }

        return attribute == null ? new ColumnInfo(name, index, pi, type.Name, pi.Name) : new ColumnInfo(name, attribute);
    }

    private static ColumnInfo GetColumnInfoByFilter(ICell header, Func<IColumnInfo, bool> columnFilter)
    {
        if (columnFilter == null) return null;

        var headerValue = GetHeaderValue(header);
        var column = new ColumnInfo(headerValue, header.ColumnIndex);

        return columnFilter(column) ? column : null;
    }

    private void LoadRowData<T>(IEnumerable<ColumnInfo> columns, IRow row, T target, RowInfo<T> rowInfo)
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

            column.RowTag = rowInfo.RowTag;
            try
            {
                var cell = row.GetCell(index);
                var propertyType = column.Attribute.PropertyUnderlyingType ?? column.Attribute.PropertyType;

                if (!MapHelper.TryGetCellValue(cell, propertyType, TrimSpaces, out object valueObj, FormulaEvaluator))
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
                    if (MapHelper.TryConvertType(valueObj, column, UseDefaultValueAttribute, out var result))
                    {
                        column.Attribute.GetSetterOrDefault(target)?.Invoke(target, result);
                    }
                    else
                    {
                        ColumnFailed(column, "Cannot convert value to the property type!");
                    }
                }
            }
            catch (Exception e)
            {
                ColumnFailed(column, e.ToString());
            }
            finally
            {
                rowInfo.RowTag = column.RowTag;
                column.RowTag = null;
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

    private static FileStream GetStreamForSave(string path) => File.Open(path, FileMode.Create, FileAccess.Write);

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
                //row.Cells.Clear();
                row.Cells?.ForEach(c => c.SetCellType(CellType.Blank)); // erase content and try keep format.
            }

            row ??= sheet.CreateRow(rowIndex);

            foreach (var column in columns)
            {
                var value = column.Attribute.GetGetterOrDefault(o)?.Invoke(o);
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
                sheet.RemoveRow(row);
                //row.Cells.Clear();
            }
            rowIndex++;
        }

        // Injects custom action for headers.
        if (overwrite && HasHeader && _headerAction != null)
        {
            firstRow?.Cells.ForEach(c => _headerAction(c));
        }
    }

    private void Save<T>(Stream stream, ISheet sheet, IEnumerable<T> objects, bool leaveOpen, bool overwrite)
    {
        Put(sheet, objects, overwrite);
        Workbook.Write(stream, leaveOpen);
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

        var attributes = Attributes.Where(p => p.Value.PropertyFullPath?.StartsWith(type.Name + ".") == true);
        var properties = new List<PropertyInfo>(type.GetProperties(MapHelper.BindingFlag).Where(p => p.PropertyType.CanBeExported()));

        // Firstly populate for those have Attribute specified.
        foreach (var pair in attributes)
        {
            var attribute = pair.Value;
            if (pair.Value.Index < 0) continue;

            var cell = row.CreateCell(attribute.Index);
            if (HasHeader)
            {
                cell.SetCellValue(attribute.Name ?? attribute.PropertyName);
            }
            properties.RemoveAll(p => p.Name == attribute.PropertyName); // Remove populated property.
        }

        var index = 0;

        // Then populate for those do not have Attribute specified.
        foreach (var pi in properties)
        {
            var key = type.Name + "." + pi.Name;
            var attribute = Attributes.TryGetValue(key, out var attribute1) ? attribute1 : null;
            if (attribute?.Ignored == true)
            {
                continue;
            }

            while (row.GetCell(index) != null)
            {
                index++;
            }

            var cell = row.CreateCell(index);
            if (HasHeader)
            {
                cell.SetCellValue(attribute?.Name ?? pi.Name);
            }
            else
            {
                new ColumnAttribute { Index = index }.SetProperty(pi, type.Name, pi.Name).MergeTo(Attributes);
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
        if (cols.TryGetValue(type, out var col))
        {
            columns = col.OfType<ColumnInfo>();
        }

        return columns?.ToList();
    }

    private void SetCell(ICell cell, object value, ColumnInfo column, bool isHeader = false, bool setStyle = true)
    {
        if (value == null || value is ICollection)
        {
            cell.SetCellValue((string)null);
        }
        else if (SkipWriteDefaultValue && !isHeader &&
                 (Equals(column.Attribute.DefaultValue, value) ||
                     (UseDefaultValueAttribute && Equals(column.Attribute.DefaultValueAttribute?.Value, value)))
                )
        {
            cell.SetCellValue((string)null);
        }
        else if (value is DateTime time)
        {
            cell.SetCellValue(time);
        }
        else if (value.GetType().IsNumeric())
        {
            cell.SetCellValue(Convert.ToDouble(value));
        }
        else if (value is bool b)
        {
            cell.SetCellValue(b);
        }
        else
        {
            cell.SetCellValue(value.ToString());
        }

        if (column != null && setStyle)
        {
            column.SetCellStyle(cell, value, isHeader, TypeFormats, _helper);
        }
    }

    #endregion Export

    private int GetFirstRowIndex(ISheet sheet) => FirstRowIndex >= 0 ? FirstRowIndex : sheet.FirstRowNum;

    #endregion
}
