# Npoi.Mapper
Convention based mapper between strong typed object and Excel data via NPOI.

## Install

Install-Package Npoi.Mapper

## Get objects from Excel (XLS or XLSX)

```C#
var stream = new FileStream("Book1.xlsx", FileMode.Open);
var importer = new Importer(stream);
var objs = importer.TakeByHeader<SampleClass>("sheet2");
```
More use cases please check out source in "test" project.

## Features

1. Import strong typed objects from Excel file (XLS or XLSX) via [NPOI](https://github.com/tonyqus/npoi)
2. No additional code required to map object properties and column headers by default naming convention (see below sectioin)
3. Explict column mapping with attributes or fluent methods
4. Support custom object factory
5. Support custom header and cell resolver
6. Support custom logic to handle adding multiple column cells into collection property.
7. Future features from you ...

## Column mapping order

1. Fluent method Map<T>
2. ColumnAttribute
3. Default naming convention (see below section)
4. MultiColumnsContainerAttribute
5. DefaultResolverType

## Default naming convention for column header mapping

1. Map column to property by name.
2. Map column to the Name of DisplayAttribute of property.
3. For column header, remove spaces, "-" and "_", and truncate from first braket ("(", "[", "{"), then map to property name

## Explicit column mapping

By fluent mapping methods:

```C#
importer.Map<SampleClass>("ColumnA", o => o.Property1)
    .Map<SampleClass>(1, o => o.Property2);
    .Ignore<SampleClass>(o => o.Property3);
    .UseLastNonBlankValue<SampleClass>(o => o.Property1);
    .DefaultResolverType = typeof (SampleColumnResolver);

var objs = importer.TakeByHeader<SampleClass>();
```

Or by Attributes tagged on object properties:

```C#
    public class SampleClass
    {
        // Other properties...
        
        [Column("ColumnABC")]
        public string Property1 { get; set; }
        
        [MultiColumnContainer(typeof(MultiColumnContainerResolver))]
        public ICollection<string> CollectionGenericProperty { get; set; }
    }
```

