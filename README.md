# Npoi.Mapper
Convention based mapper between strong typed object and Excel data via NPOI.  
This project comes up with a task of my work, I am using it a lot in my project. Feel free to file bugs or raise pull requests...

## Install from NuGet
In the Package Manager Console:

`PM> Install-Package Npoi.Mapper`

## Get objects from Excel (XLS or XLSX)

```C#
var mapper = new Mapper("Book1.xlsx");
var objs1 = mapper.Take<SampleClass>("sheet2");

// You can take objects from the same sheet with different type.
var objs2 = mapper.Take<AnotherClass>("sheet2");
```
More use cases please check out source in "test" project.

## Export objects to Excel (XLS or XLSX)

### 1. Export objects.
Set **`overwrite`** parameter to false to use existing columns and formats, otherwise always create new file.
```C#
//var objects = ...
var mapper = new Mapper();
mapper.Save("test.xlsx",  objects, "newSheet", overwrite: false);
```

### 2. Export tracked objects.
Set **`TrackObjects`** property to true, objects can be tracked after a `Take` method and then you can modify and save them back.
```C#
var mapper = new Mapper("Book1.xlsx");
mapper.TrackObjects = true; // It's default true.
var objectInfos = mapper.Take<SampleClass>("sheet2"); // You can Take first then modify tracked objects.
var objectsDict = mapper.Objects; // Also you can directly access objects in a sheet by property.
mapper.Save("test.xlsx",  "sheet2");
```

### 3. Put different types of objects into memory workbook and export together.
Set **`overwrite`** parameter to true, existing data rows will be overwritten, otherwise new rows will be appended.
```C#
var mapper = new Mapper("Book1.xlsx");
mapper.Put(products, "sheet1", true);
mapper.Put(orders, "sheet2", false);
mapper.Save("Book1.xlsx");
```

## Features

1. Import POCOs from Excel file (XLS or XLSX) via [NPOI](https://github.com/tonyqus/npoi)
2. Export objects to Excel file (XLS or XLSX) (inspired by [ExcelMapper](https://github.com/mganss/ExcelMapper))
3. No code required to map object properties and column headers by default naming convention (see below sectioin)
4. Support to escape and truncate chars in column header for mapping
5. Also support explicit column mapping with attributes or fluent methods
6. Support custom object factory injection
7. Support custom header and cell resolver
8. Support custom logic to handle multiple columns for collection property
9. Support Excel built-in format and custom format for exporting (see Column format section)

## Column mapping order

1. Fluent method `Map<T>`
2. `ColumnAttribute`
3. Default naming convention (see below section)
4. `Mapper.DefaultResolverType` property

## Default naming convention for column header mapping

1. Map column to property by name.
2. Map column to the Name of `DisplayAttribute` of property.
3. For column header, ignore non-alphabetical chars ("-", "_", "|' etc.), and truncate from first braket ("(", "[", "{"), then map to property name. Ignored chars and truncation chars can be customized.

## Explicit column mapping

By fluent mapping methods:

```C#
mapper.Map<SampleClass>("ColumnA", o => o.Property1)
    .Map<SampleClass>(1, o => o.Property2)
    .Ignore<SampleClass>(o => o.Property3)
    .UseLastNonBlankValue<SampleClass>(o => o.Property1)
    .Format<SampleClass>("yyyy/MM/dd", o => o.DateProperty)
    .DefaultResolverType = typeof (SampleColumnResolver);
```

Or by Attributes tagged on object properties:

```C#
    public class SampleClass
    {
        // Other properties...
        
        [Display(Name = "Display Name")]
        public string DisplayNameProperty { get; set; }
        
        [Column(1)]
        public string Property1 { get; set; }
        
        [Column("ColumnABC")]
        public string Property2 { get; set; }
        
        [Column(BuiltinFormat = 0xf)]
        public DateTime BuiltinFormatProperty { get; set; }
        
        [Column(CustomFormat = "0%")]
        public double CustomFormatProperty { get; set; }
        
        [Column(ResolverType = typeof(MultiColumnContainerResolver))]
        public ICollection<string> CollectionGenericProperty { get; set; }
        
        [UseLastNonBlankValue]
        public string UseLastNonBlankValueAttributeProperty { get; set; }
        
        [Ignore]
        public string IgnoredAttributeProperty { get; set; }
    }
```

## Column format

By method:

```C#
mapper.Format<SampleClass>("yyyy/MM/dd", o => o.DateProperty)
    .Format<SampleClass>("0%", o => o.DoubleProperty);
```

Or by `ColumnAttribute`

```C#
    public class SampleClass
    {
        [Column(BuiltinFormat = 0xf)]
        public DateTime BuiltinFormatProperty { get; set; }
        
        [Column(CustomFormat = "0%")]
        public double CustomFormatProperty { get; set; }
    }
```

You can use both **[builtin formats](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html)** and **[custom formats](https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4)**.

## Custom column resolver
Implement **`IColumnResolver`** to handle complex scenarios. Such as data conversion or retrieve values cross columns for a collection property.

```C#
    public class MultiColumnContainerResolver : IColumnResolver<SampleClass>
    {
        public bool IsColumnMapped(ref object headerValue, int index)
        {
            try
            {
                // Custom logic to determine whether or not to map and include this column.
                // Header value is either in string or double. Try convert by needs.
                if (index > 30 && index <= 40 && headerValue is double)
                {
                    // Assign back header value and use it from TryResolveCell method.
                    headerValue = DateTime.FromOADate((double)headerValue);

                    return true;
                }
            }
            catch
            {
                // Does nothing here and return false eventually.
            }

            return false;
        }

        public bool TryTakeCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            if (!(columnInfo.HeaderValue is DateTime)) return false;

            // Custom logic to handle the cell value.
            target.CollectionGenericProperty.Add(((DateTime)columnInfo.HeaderValue).ToLongDateString() + cellValue);

            return true;
        }

        public bool TryPutCell(ColumnInfo<SampleClass> columnInfo, out object cellValue, SampleClass source)
        {
            cellValue = null;

            // Note: return false to indicate a failure; and that will increase error count.
            if (!(columnInfo?.HeaderValue is DateTime)) return false;

            var s = ((DateTime)columnInfo.HeaderValue).ToLongDateString();

            // Custom logic to set the cell value.
            if (source.CollectionGenericProperty.Count > 0 && columnInfo.Attribute.Index == 31)
            {
                cellValue = source.CollectionGenericProperty.ToList()[0].Remove(0, s.Length);
            }
            else if (source.CollectionGenericProperty.Count > 1 && columnInfo.Attribute.Index == 33)
            {
                cellValue = source.CollectionGenericProperty.ToList()[1].Remove(0, s.Length);
            }

            return true;
        }
    }
```

## Change log

### v2.0.7
* Fixed issue #3: **`Put`** method does not work when using a custom column resolver.

### v2.0.6
* Fixed issue #1: cannot ignore properties from base class.

### v2.0.5
* Convert **`ColumnResolver`** to **`IColumnResolver`** interface to inject custom logic when export data to file.

### v2.0.4
* Added **`Put`** methods and new **`Save`** methods, so you can put different type of objects in memory workbook first and then save them together.

### v2.0.3
* Support **`overwrite`** parameter for exporting data, use existing file if set to false.

