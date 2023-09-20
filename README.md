# Npoi.Mapper

[![NuGet Version](https://shields.io/nuget/v/npoi.mapper?style=flat-square&logo=nuget)](https://www.nuget.org/packages/Npoi.Mapper)
[![NuGet Downloads](https://shields.io/nuget/dt/npoi.mapper?style=flat-square&logo=nuget)](https://www.nuget.org/packages/Npoi.Mapper)
[![Open Issues](https://shields.io/github/issues/donnytian/npoi.mapper?style=flat-square&logo=github)](https://github.com/donnytian/Npoi.Mapper/issues)
[![Open PRs](https://shields.io/github/issues-pr/donnytian/npoi.mapper?style=flat-square&logo=github)](https://github.com/donnytian/Npoi.Mapper/pulls)

| Branch  | Build & Test                                                                                                                | Publish                                                                                                    |
|:--------|:----------------------------------------------------------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------|
| master  | ![master build status](https://github.com/donnytian/Npoi.Mapper/actions/workflows/build-test.yml/badge.svg?branch=master)   | ![ Publish Status](https://github.com/donnytian/Npoi.Mapper/actions/workflows/publish-nuget.yml/badge.svg) |
| develop | ![develop build status](https://github.com/donnytian/Npoi.Mapper/actions/workflows/build-test.yml/badge.svg?branch=develop) | N/A                                                                                                        |


Convention-based mapper between strong typed object and Excel data via NPOI (Telegram group https://t.me/npoidevs).  
This project comes up with a task of my work, I am using it a lot in my project. Feel free to file bugs or raise pull requests...

Support to import and export as *`dynamic`* type.

Support **nested property mapping** (only work for explicit mapping by `Map` method).

## [Change log](Changelog.md)
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
DateTime date = obj3[0].Value.DateColumn;
double number = obj3[0].Value.NumberColumn;
string text = obj3[0].Value.AC; // If the column doesn't have a header name, Excel display name like "AC" will be populated.
```

### Differences between `Take<dynamic>()` and `TakeDynamicWithColumnType()`
`Take<Dynamic>` is a quick way to take data by conventions, all column types will be auto-detected from the first data row.

However, we may encounter unexpected data at the first row in the real world, that makes the mapper to infer a wrong type.
For example, we actually want a `string` column, but a numeric value sit in the first row cell, this makes mapper take this column as `double` and report errors for the following non-numeric values. 

`TakeDynamicWithColumnType()` allow you predefine the type for any column by accepting an extra parameter.
```csharp
var mapper = new Mapper(workbook);
var objs = mapper.TakeDynamicWithColumnType(header =>
    header.ColumnIndex switch    // Inspect column index or header cell to make decision.
    {
        0 => typeof(int),       // Make the 1st column as int
        1 => typeof(DateTime),  // Make the 2nd column as DateTime
        2 => typeof(string),    // Make the 3rd column as string
        _ => null,              // return null to let mapper detect from the first data row.
    });


// Or simply take all columns as string.
var objs = mapper.TakeDynamicWithColumnType(_ => typeof(string));
```

### Import numeric cell as `string`
If the target property is in the type of `string`, cell format and formula will be respected. That means you will get exact same string as what is displayed in Excel instead of underlying number or formula.

More use cases please check out source in "test" project.

## Export objects to Excel (XLS or XLSX)

### 1. Export objects.
Set **`overwrite`** parameter to false to use existing columns and formats, otherwise always create new file.
```C#
//var objects = ...
var mapper = new Mapper();
mapper.Save("test.xlsx",  objects, "newSheet", leaveOpen: false, overwrite: false);
```

### 2. ~~Export tracked objects.~~ (This feature was removed from 6.2)
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
mapper.Save("Book1.xlsx", leaveOpen: false);
```

## Features

1. Import POCOs from Excel file (XLS or XLSX) via [NPOI](https://github.com/tonyqus/npoi)
2. Export objects to Excel file (XLS or XLSX) (inspired by [ExcelMapper](https://github.com/mganss/ExcelMapper))
3. No code required to map object properties and column headers by default naming convention (see below section)
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
3. For column header, ignore non-alphabetical chars ("-", "_", "|' etc.), and truncate from first bracket ("(", "[", "{"), then map to property name. Ignored chars and truncation chars can be customized.

## Explicit column mapping

By fluent mapping methods.

Nested properties are supported.
For the import, if any property in the middle of the chain is null, the mapper will create new instance -
just make sure it has a public parameterless constructor. 

```C#
mapper.Map<SampleClass>("ColumnA", o => o.Property1)
    .Map<SampleClass>(1, o => o.Property2)
    .Map<SampleClass>(2, o => o.Customer.Address.ZipCode)
    .Ignore<SampleClass>(o => o.Property3)
    .UseLastNonBlankValue<SampleClass>(o => o.Property1)
    .Format<SampleClass>("yyyy/MM/dd", o => o.DateProperty)
    .DefaultResolverType = typeof(SampleColumnResolver);
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
                    column.RowTag = "MyCustomObject"; // this can be retrieved in the result items.
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
