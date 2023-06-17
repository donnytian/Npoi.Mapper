using System;
using System.Linq;
using Npoi.Mapper;
using NUnit.Framework;
using test.Sample;

namespace test
{
    [TestFixture]
    public class CustomResolverTests : TestBase
    {
        [Test]
        public void SingleColumnResolverTest()
        {
            // Arrange
            var date1 = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            // We will import column with index of 51.
            sheet.GetRow(0).CreateCell(51).SetCellValue(date1);
            sheet.GetRow(1).CreateCell(51).SetCellValue(str1);

            // Act "Take"
            var mapper = new Mapper(workbook);
            mapper.Map<SampleClass>(51, o => o.SingleColumnResolverProperty,
                (column, target) => // tryTake resolver : Custom logic to take cell value into target object.
                {
                    // Note: return false to indicate a failure; and that will increase error count.
                    if (column.HeaderValue == null || column.CurrentValue == null) return false;

                    if (column.HeaderValue is double)
                    {
                        column.HeaderValue = DateTime.FromOADate((double)column.HeaderValue);
                    }

                    // Custom logic to get the cell value.
                    ((SampleClass)target).SingleColumnResolverProperty = ((DateTime)column.HeaderValue).ToLongDateString() + column.CurrentValue;

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
                    column.CurrentValue = ((SampleClass)source).SingleColumnResolverProperty?.Remove(0, s.Length);

                    return true;
                }
                );

            var objs = mapper.Take<SampleClass>().ToList();

            // Assert "Take"
            Assert.IsNotNull(objs);
            Assert.AreEqual(date1.ToLongDateString() + str1, objs[0].Value.SingleColumnResolverProperty);

            // Act "Put"
            objs[0].Value.SingleColumnResolverProperty = date1.ToLongDateString() + str2;
            mapper.Put(new[] { objs[0].Value });

            // Assert "Put"
            Assert.AreEqual(str2, sheet.GetRow(1).GetCell(51).StringCellValue);
        }

        [Test]
        public void MultiColumnContainerTest()
        {
            // Arrange
            var date1 = DateTime.Now;
            var date2 = date1.AddMonths(1);
            const string str1 = "aBC";
            const string str2 = "BCD";
            const string str3 = "_PutTest";
            var workbook = GetSimpleWorkbook(date1, str1);

            // We will import columns with index of 31 and 33 into a collection property.
            workbook.GetSheetAt(1).GetRow(0).CreateCell(31).SetCellValue(date1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(33).SetCellValue(date2);

            workbook.GetSheetAt(1).GetRow(1).CreateCell(31).SetCellValue(str1);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(33).SetCellValue(str2);

            // Act
            var mapper = new Mapper(workbook);
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

            // Act Take
            var objs = mapper.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];

            Assert.AreEqual(2, obj.Value.CollectionGenericProperty.Count);

            var list = obj.Value.CollectionGenericProperty.ToList();

            // Assert Take
            Assert.AreEqual(date1.ToLongDateString() + str1, list[0]);
            Assert.AreEqual(date2.ToLongDateString() + str2, list[1]);

            // Act Put
            obj.Value.CollectionGenericProperty.Clear();
            obj.Value.CollectionGenericProperty.Add(date1.ToLongDateString() + str3);
            obj.Value.CollectionGenericProperty.Add(date2.ToLongDateString() + str3);
            mapper.Put(new[] { objs[0].Value }, 1);

            // Assert "Put"
            var sheet = workbook.GetSheetAt(1);
            Assert.AreEqual(str3, sheet.GetRow(1).GetCell(31).StringCellValue);
            Assert.AreEqual(str3, sheet.GetRow(1).GetCell(33).StringCellValue);
        }

        //https://github.com/donnytian/Npoi.Mapper/issues/23
        [Test]
        public void WithInvalidEnum_TryTake_ShouldBeCalled()
        {
            // Arrange
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            // Header row
            sheet.GetRow(0).CreateCell(0).SetCellValue("EnumProperty");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(11); // Invalid enum value.

            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(0, o => o.EnumProperty, (column, obj) =>
            {
                ((SampleClass) obj).EnumProperty = SampleEnum.Value3;
                return true;
            }, null);
            var items = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.AreEqual(SampleEnum.Value3, items[0].Value.EnumProperty);
        }

        //https://github.com/donnytian/Npoi.Mapper/issues/64
        [Test]
        public void Map_TryTakeDateTimeFromDouble_ShouldGetDateTime()
        {
            // Arrange
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var value = DateTime.Now;
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            var columnName = nameof(SampleClass.DateProperty);
            const string format = "m/d/yyyy h:mm";

            // Header row
            sheet.GetRow(0).CreateCell(0).SetCellValue(columnName);

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(value.ToString(format));

            var mapper = new Mapper(workbook);
            mapper.UseFormat(typeof(DateTime), format);

            // Act
            var items = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.AreEqual(value.ToString(format), items[0].Value.DateProperty.ToString(format));
        }
        
        [Test]
        public void RowTagForTryTakeTest()
        {
            // Arrange
            var workbook = GetSimpleWorkbook(DateTime.Now, "foo");
            var sheet = workbook.GetSheetAt(1);
            var row2 = sheet.CreateRow(2);
            row2.CreateCell(1).SetCellValue("bar");
            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(1, o => o.StringProperty, (column, obj) =>
            {
                ((SampleClass)obj).StringProperty = column.CurrentValue as string;
                column.RowTag ??= column.CurrentValue;
                return true;
            }, null);
            var items = mapper.Take<SampleClass>(1).ToList();

            // Assert
            Assert.AreEqual(2, items.Count);
            Assert.AreEqual("foo", items[0].RowTag);
            Assert.AreEqual("bar", items[1].RowTag);
        }
    }
}
