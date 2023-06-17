# Change log

## v6
* **Breaking change:** Target on net6.0 only; sorry for net45
* **Breaking change:** Upgrade NPOI version to 2.6.0, so a new `leaveOpen` parameter is required for any `Save()` method
* **Breaking change:** `Save()` methods now will try to open existing file instead of create new one; thus the `overwrite` parameter will overwrite rows in existing file as expected.
* The default value of `maxErrorRows` in `Take()` methods is change to 100 from 10.
* New property `RowTag` in the result items to associate custom object during `TryTake()` resolver.
* New setting `SkipHiddenRows` and some minor refines.

## From v4 go on, we will use Github Release feature to track change logs.

## v4
* Upgrade NPOI to latest 2.5.6
* **Breaking change**: Removed support for NET4.0 since the latest NPOI does not support it.
* **Breaking change**: Make assembly strong-named - #80
* New option `SkipBlankRows` and `TrimSpaces` - #83
* Support default value if data source is null or empty - #91
* Thanks @POFerro for above feature enhancements :)

## v3.5.1
* Include exception details in `rowInfo.ErrorMessage` for data import.

## v3.5
* Support .NETFramework 4.0, .NETFramework 4.5 and .NETStandard 2.0 in order to align with NPOI.
* Support Guid type
* When map a column by header name, apply trim on both Excel column name and mapped name.

## v3.4.1
* Workaround to fix a NPOI regression for `sheet.RemoveRow()`. The issue was fixed on NPOI 2.4.1 but has returned on 2.5.1.

## v3.4
* Removed ~~`HeaderRowIndex`~~ property. Replaced it by `FirstRowIndex`.

## v3.3
* Upgrade dependency of NPOI to 2.4.1
* Support .NETStandard 2.0

## v3.2
* Fixed issue #24: Added a Map overload for export to specify both column index and name
* Fixed issue #25: ForHeader action now will be executed after data export, so Sheet.AutoSizeColumn should work properly

## v3.1.1
* Fixed issue #22: ForHeader method will not executed in certain case
* Fixed issue #23: Cannot apply TryTake if there is error when parsing Enum

## v3.1
* Added overload to ignore properties by string
* Added method 'ForHeader' to allow set header's cell style when exporting.
* A few bug fixes for Ignore and Put

## v3.0.2
* Support .NET Framework 4.5.

## v3.0.1
* Fixed issue #7: **`IgnoredNameChars`** not working properly.

## v3.0
* New feature: Use **`Take<dynamic>`** so you don't even need a predefined type to import data, mapper will do it for you.
* Breaking change: Removed support for **`BuiltinFormat`**, please use **`CustomFormat`** only.
* Breaking change: Removed support for **`IColumnResolver`**, instead use overloads of **`Map`** method to specify custom resolvers.

## v2.1.1
* Fixed issue #5: **`UseFormat`** does not support for nullable types when data in first row is null.
* Fixed issue #6: Added a overload method for **`Map`** to map a property by a string.

## v2.1
* Enhancement for #4: Added **`UseFormat`** method to use a default format for all properties that have a same type.
* Support Nullable properties.
* Builtin format will be obsolete in v3.0.
 
## v2.0.7
* Fixed issue #3: **`Put`** method does not work when using a custom column resolver.

## v2.0.6
* Fixed issue #1: cannot ignore properties from base class.

## v2.0.5
* Convert **`ColumnResolver`** to **`IColumnResolver`** interface to inject custom logic when export data to file.

## v2.0.4
* Added **`Put`** methods and new **`Save`** methods, so you can put different type of objects in memory workbook first and then save them together.

## v2.0.3
* Support **`overwrite`** parameter for exporting data, use existing file if set to false.

