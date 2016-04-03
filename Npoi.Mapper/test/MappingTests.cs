using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;
using test.Sample;

namespace test
{
    /// <summary>
    /// Column mapping tests.
    /// </summary>
    [TestClass]
    public class MappingTests : TestBase
    {
        [TestMethod]
        public void ColumnIndexTest()
        {
            // Prepare
            const string str = "aBC";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(11);

            sheet.GetRow(0).CreateCell(11).SetCellValue("targetColumn");
            sheet.GetRow(11).CreateCell(11).SetCellValue(str);

            var importer = new Mapper(workbook);

            // Act
            importer.Map<SampleClass>(11, o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.GeneralProperty);
        }

        [TestMethod]
        public void ColumnNameTest()
        {
            // Prepare
            const string str = "aBC";
            const string name = "targetColumn";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(11);

            sheet.GetRow(0).CreateCell(11).SetCellValue(name);
            sheet.GetRow(11).CreateCell(11).SetCellValue(str);

            var importer = new Mapper(workbook);

            // Act
            importer.Map<SampleClass>(name, o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.GeneralProperty);
        }

        [TestMethod]
        public void ColumnsWithSameNameTest()
        {
            // Prepare
            const string str1 = "aBC";
            const string str2 = "aBC";
            const string name = "targetColumn";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(11);

            sheet.GetRow(0).CreateCell(7).SetCellValue(name);
            sheet.GetRow(0).CreateCell(9).SetCellValue(name);

            sheet.GetRow(11).CreateCell(7).SetCellValue(str1);
            sheet.GetRow(11).CreateCell(9).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            importer.Map<SampleClass>(name, o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>().ToList();

            // Assert
            var obj = objs[0];
            Assert.AreEqual(str2, obj.Value.GeneralProperty);
        }

        [TestMethod]
        public void DefaultResolverTypeTest()
        {
            // Prepare
            var date1 = DateTime.Now;
            var date2 = date1.AddMonths(1);
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            sheet.GetRow(0).CreateCell(41).SetCellValue(date1);
            sheet.GetRow(0).CreateCell(43).SetCellValue(date2);

            sheet.GetRow(1).CreateCell(41).SetCellValue(str1);
            sheet.GetRow(1).CreateCell(43).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            importer.DefaultResolverType = typeof(DefaultColumnResolver);
            var objs = importer.Take<SampleClass>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];

            Assert.AreEqual(2, obj.Value.GeneralCollectionProperty.Count);

            var list = obj.Value.GeneralCollectionProperty.ToList();

            Assert.AreEqual(date1.ToLongDateString() + str1, list[0]);
            Assert.AreEqual(date2.ToLongDateString() + str2, list[1]);
        }

        [TestMethod]
        public void IgnoredTest()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetSimpleWorkbook(date, str1);

            var header = workbook.GetSheetAt(1).GetRow(0).CreateCell(41);
            header.SetCellValue(nameof(sample.GeneralProperty));
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);

            var importer = new Mapper(workbook);

            // Act
            importer.Ignore<SampleClass>(o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNull(objs[0].Value.GeneralProperty);
        }

        [TestMethod]
        public void UseLastNonBlankValueTest()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetSimpleWorkbook(date, str1);

            var header = workbook.GetSheetAt(1).GetRow(0).CreateCell(41);
            header.SetCellValue(nameof(sample.GeneralProperty));

            // Create 4 rows, row 22 and 23 have empty values.
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);
            workbook.GetSheetAt(1).CreateRow(22).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(23).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(24).CreateCell(41).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            importer.UseLastNonBlankValue<SampleClass>(o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            var obj = objs[1];
            Assert.AreEqual(str1, obj.Value.GeneralProperty);

            obj = objs[2];
            Assert.AreEqual(str1, obj.Value.GeneralProperty);

            obj = objs[3];
            Assert.AreEqual(str1, obj.Value.GeneralProperty);

            obj = objs[4];
            Assert.AreEqual(str2, obj.Value.GeneralProperty);
        }

        [TestMethod]
        public void MethodOverAttributeTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            const string str3 = "EFG";
            var workbook = GetSimpleWorkbook(date, str1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11).SetCellValue("ColumnIndexAttributeProperty");
            workbook.GetSheetAt(1).GetRow(0).CreateCell(12).SetCellValue("targetColumn");
            workbook.GetSheetAt(1).GetRow(0).CreateCell(13).SetCellValue("By Name");

            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str1);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(12).SetCellValue(str2);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(13).SetCellValue(str3);

            var importer = new Mapper(workbook);

            // Act
            importer.Map<SampleClass>("targetColumn", o => o.ColumnIndexAttributeProperty);
            importer.Map<SampleClass>(13, o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(str2, objs[0].Value.ColumnIndexAttributeProperty);
            Assert.AreEqual(str3, objs[0].Value.GeneralProperty);
            Assert.IsNull(objs[0].Value.ColumnNameAttributeProperty);
        }

        [TestMethod]
        public void NameOverIndexTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            const string str3 = "EFG";
            const string str4 = "FGH";
            var workbook = GetSimpleWorkbook(date, str1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(12).SetCellValue("ColumnIndexAttributeProperty");
            workbook.GetSheetAt(1).GetRow(0).CreateCell(13);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(14).SetCellValue("targetColumn");

            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str1);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(12).SetCellValue(str2);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(13).SetCellValue(str3);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(14).SetCellValue(str4);

            var importer = new Mapper(workbook);

            // Act
            importer.Map<SampleClass>(13, o => o.ColumnIndexAttributeProperty);
            importer.Map<SampleClass>("targetColumn", o => o.ColumnIndexAttributeProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(str4, objs[0].Value.ColumnIndexAttributeProperty);
        }
    }
}
