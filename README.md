# Npoi.Mapper
Convention based mapper between strong typed object and Excel data via NPOI.

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
3. Explict column mapping by attributes such as by column name and index
4. Support custom object factory
5. Support custom header and cell resolver
6. Support custom logic to handle adding multiple column cells into collection property.
7. Future features from you ...

## Column mapping order

1. ColumnNameAttribute
2. ColumnIndexAttribute
3. Default naming convention (see below section)
4. MultiColumnsContainerAttribute

## Default nameing convention for column header mapping

1. Map column to property by name.
2. Map column to the Name of DisplayAttribute of property.
3. For column header, remove spaces, "-" and "_", and truncate from first braket ("(", "[", "{"), then map to property name
