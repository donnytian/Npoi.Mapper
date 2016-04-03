# Npoi.Mapper
Convention based mapper between strong typed object and Excel data via NPOI.  
Feel free to file bugs or raise pull requests...

## Install

PM> Install-Package Npoi.Mapper

## Get objects from Excel (XLS or XLSX)

```C#
var mapper = new Mapper("Book1.xlsx");
var objs1 = mapper.Take<SampleClass>("sheet2");

// You can take objects from the same sheet with different type.
var objs2 = mapper.Take<AnotherClass>("sheet2");
```
More use cases please check out source in "test" project.

## Export objects to Excel (XLS or XLSX)

```C#
// Export objects.
//var objects = ...
var mapper = new Mapper();
mapper.Save("test.xlsx",  objects, "newSheet");

// Export tracked objects.
var mapper = new Mapper("Book1.xlsx");
var objectInfos = mapper.Take<SampleClass>("sheet2").ToList();
// Modify tracked objects...then save back..
mapper.Save("test.xlsx",  "sheet2");
mapper.Save("test.xlsx",  objectInfos.Select(info => info.Value), "sheet3");
```

## Features

1. Import POCOs from Excel file (XLS or XLSX) via [NPOI](https://github.com/tonyqus/npoi)
2. Export objects to Excel file (XLS or XLSX) (inspired by [ExcelMapper](https://github.com/mganss/ExcelMapper))
3. No code required to map object properties and column headers by default naming convention (see below sectioin)
4. Support escaping and truncate chars in column header for mapping
4. Also support explicit column mapping with attributes or fluent methods
5. Support built-in and custom Excel cell format
6. Support custom object factory injection
7. Support custom header and cell resolver
8. Support custom logic to handle multiple column for collection property.

## Column mapping order

1. Fluent method Map<T>
2. ColumnAttribute
3. Default naming convention (see below section)
4. DefaultResolverType

## Default naming convention for column header mapping

1. Map column to property by name.
2. Map column to the Name of DisplayAttribute of property.
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

