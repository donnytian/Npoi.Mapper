# Npoi.Mapper
Convention-based mapper between strong typed object and Excel data via NPOI.  
This project comes up with a task of my work, I am using it a lot in my project. Feel free to file bugs or raise pull requests...

<font color=brown>v3 now support to import and export as **`dynamic`** type.</font>
## Install from NuGet
In the Package Manager Console:

`PM> Install-Package Npoi.Mapper`

## Get strong-typed objects from Excel (XLS or XLSX)

```C#
var mapper = new Mapper("Book1.xlsx");
var objs1 = mapper.Take<SampleClass>("sheet2");

// You can take objects from the same sheet with different type.
var objs2 = mapper.Take<AnotherClass>("sheet2");

// Even you can use dynamic type.
// DateTime, double and string will be auto-detected for object properties.
// You will get a DateTime property only if the cell in Excel was formatted as a date, otherwise it will be a double.
var objs3 = mapper.Take<dynamic>("sheet1").ToList();
DateTime date = obj3[0].DateColumn;
double number = obj3[0].NumberColumn;
string text = obj3[0].AC; // If the column doesn't have a header name, Excel display name like "AC" will be populated.
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

### 3. Put different types of objects into memory workbook then export together.
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
9. Support custom format for exporting (see Column format section)

## Column mapping order

1. Fluent method `Map<T>`
2. `ColumnAttribute`
3. Default naming convention (see below section)

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
        
        [Column(CustomFormat = "0%")]
        public double CustomFormatProperty { get; set; }
        
        [UseLastNonBlankValue]
        public string UseLastNonBlankValueAttributeProperty { get; set; }
        
        [Ignore]
        public string IgnoredAttributeProperty { get; set; }
    }
```

## Column format

When you use a format during import, it will try to parse string value with specified format.

When you use a format during export, it will try to set Excel display format with specified format.

By method:

```C#
    mapper.Format<SampleClass>("yyyy/MM/dd", o => o.DateProperty)
          .Format<SampleClass>("0%", o => o.DoubleProperty);
```

Or by `ColumnAttribute`:

```C#
    public class SampleClass
    {
        [Column(CustomFormat = "yyyy-MM-dd")]
        public DateTime DateTimeFormatProperty { get; set; }
        
        [Column(CustomFormat = "0%")]
        public double CustomFormatProperty { get; set; }
    }
```

Or if you want to set format for all properties in a same type:

```C#
    mapper.UseFormat(typeof(DateTime), "yyyy.MM.dd hh.mm.ss");
```
You can find format details at **[custom formats](https://support.office.com/en-us/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4)**.

## Custom column resolver
Use overload of **`Map`** method to handle complex scenarios. Such as data conversion or retrieve values cross columns for a collection property.

```C#
    mapper.Map(
                column => // column filter : Custom logic to determine whether or not to map and include an unmapped column.
                {
                    // Header value is either in string or double. Try convert by needs.

                    var index = column.Attribute.Index;

                    if ((index == 31 || index == 33) && column.HeaderValue is double)
                    {
                        // Assign back header value and use it from TryTake method.
                        column.HeaderValue = DateTime.FromOADate((double)column.HeaderValue);

                        return true;
                    }

                    return false;
                },
                (column, target) => // tryTake resolver : Custom logic to take cell value into target object.
                {
                    // Note: return false to indicate a failure; and that will increase error count.
                    if (column.HeaderValue == null || column.CurrentValue == null) return false;
                    if (!(column.HeaderValue is DateTime)) return false;

                    ((SampleClass)target).CollectionGenericProperty.Add(((DateTime)column.HeaderValue).ToLongDateString() + column.CurrentValue);

                    return true;
                },
                (column, source) => // tryPut resolver : Custom logic to put property value into cell.
                {
                    if (column.HeaderValue is double)
                    {
                        column.HeaderValue = DateTime.FromOADate((double)column.HeaderValue);
                    }

                    var s = ((DateTime)column.HeaderValue).ToLongDateString();

                    // Custom logic to set the cell value.
                    var sample = (SampleClass) source;
                    if (column.Attribute.Index == 31 && sample.CollectionGenericProperty.Count > 0)
                    {
                        column.CurrentValue = sample.CollectionGenericProperty?.ToList()[0].Remove(0, s.Length);
                    }
                    else if (column.Attribute.Index == 33 && sample.CollectionGenericProperty.Count > 1)
                    {
                        column.CurrentValue = sample.CollectionGenericProperty?.ToList()[1].Remove(0, s.Length);
                    }

                    return true;
                }
                );
```

## Change log

### v3.0
* New feature: Use **`Take<dynamic>`** so you don't even need a predefined type to import data, mapper will do it for you.
* Breaking change: Removed support for **`BuiltinFormat`**, please use **`CustomFormat`** only.
* Breaking change: Removed support for **`IColumnResolver`**, instead use overloads of **`Map`** method to specify custom resolvers.

### v2.1.1
* Fixed issue #5: **`UseFormat`** does not support for nullable types when data in first row is null.
* Fixed issue #6: Added a overload method for **`Map`** to map a property by a string.

### v2.1
* Enhancement for #4: Added **`UseFormat`** method to use a default format for all properties that have a same type.
* Support Nullable properties.
* Builtin format will be obsolete in v3.0.
 
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

