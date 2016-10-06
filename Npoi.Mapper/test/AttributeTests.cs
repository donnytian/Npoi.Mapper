using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;

using test.Sample;

namespace test
{
    [TestClass]
    public class AttributeTests : TestBase
    {
        [TestMethod]
        public void ColumnAttributeIndexTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11).SetCellValue("targetColumn");
            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.ColumnIndexAttributeProperty);
        }

        [TestMethod]
        public void ColumnAttributeNameTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(21).SetCellValue("By Name");
            workbook.GetSheetAt(1).GetRow(1).CreateCell(21).SetCellValue(str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.ColumnNameAttributeProperty);
        }

        [TestMethod]
        public void DisplayNameTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(21).SetCellValue("Display Name");
            workbook.GetSheetAt(1).GetRow(1).CreateCell(21).SetCellValue(str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.DisplayNameProperty);
        }

        [TestMethod]
        public void SingleColumnResolverTypeTest()
        {
            // Prepare
            var date1 = DateTime.Now;
            //var date2 = date1.AddMonths(1);
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            sheet.GetRow(0).CreateCell(51).SetCellValue(date1);
            //sheet.GetRow(0).CreateCell(53).SetCellValue(date2);

            sheet.GetRow(1).CreateCell(51).SetCellValue(str1);
            //sheet.GetRow(1).CreateCell(53).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act "Take"
            var objs = importer.Take<SampleClass>().ToList();

            // Assert "Take"
            Assert.IsNotNull(objs);
            Assert.AreEqual(date1.ToLongDateString() + str1, objs[0].Value.SingleColumnResolverProperty);

            // Act "Put"
            objs[0].Value.SingleColumnResolverProperty = date1.ToLongDateString() + str2;
            importer.Put(new[] { objs[0].Value });

            // Assert "Put"
            Assert.AreEqual(str2, sheet.GetRow(1).GetCell(51).StringCellValue);
        }

        [TestMethod]
        public void MultiColumnContainerTest()
        {
            // Prepare
            var date1 = DateTime.Now;
            var date2 = date1.AddMonths(1);
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetSimpleWorkbook(date1, str1);

            workbook.GetSheetAt(1).GetRow(0).CreateCell(31).SetCellValue(date1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(33).SetCellValue(date2);

            workbook.GetSheetAt(1).GetRow(1).CreateCell(31).SetCellValue(str1);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(33).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];

            Assert.AreEqual(2, obj.Value.CollectionGenericProperty.Count);

            var list = obj.Value.CollectionGenericProperty.ToList();

            Assert.AreEqual(date1.ToLongDateString() + str1, list[0]);
            Assert.AreEqual(date2.ToLongDateString() + str2, list[1]);
        }

        [TestMethod]
        public void UseLastNonBlankValueAttributeTest()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetSimpleWorkbook(date, str1);

            var header = workbook.GetSheetAt(1).GetRow(0).CreateCell(41);
            header.SetCellValue(nameof(sample.UseLastNonBlankValueAttributeProperty));

            // Create 4 rows, row 22 and 23 have empty values.
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);
            workbook.GetSheetAt(1).CreateRow(22).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(23).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(24).CreateCell(41).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(5, objs.Count);

            var obj = objs[1];
            Assert.AreEqual(str1, obj.Value.UseLastNonBlankValueAttributeProperty);

            obj = objs[2];
            Assert.AreEqual(str1, obj.Value.UseLastNonBlankValueAttributeProperty);

            obj = objs[3];
            Assert.AreEqual(str1, obj.Value.UseLastNonBlankValueAttributeProperty);

            obj = objs[4];
            Assert.AreEqual(str2, obj.Value.UseLastNonBlankValueAttributeProperty);
        }

        [TestMethod]
        public void IgnoreAttributeTest()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetSimpleWorkbook(date, str1);

            workbook.GetSheetAt(1).GetRow(0).CreateCell(41).SetCellValue(nameof(sample.IgnoredAttributeProperty));
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);

            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNull(objs[0].Value.IgnoredAttributeProperty);
        }
    }
}
