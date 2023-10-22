using System;
using System.Linq;
using System.Reflection;
using Npoi.Mapper;
using NUnit.Framework;
using test.Sample;

namespace test
{
    /// <summary>
    /// Column mapping tests.
    /// </summary>
    [TestFixture]
    public class MappingTests : TestBase
    {
        [Test]
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

        [Test]
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

        [Test]
        public void ColumnName_MapPropertyByString()
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
            importer.Map<SampleClass>(name, "GeneralProperty");
            var objs = importer.Take<SampleClass>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.GeneralProperty);
        }

        [Test]
        public void ColumnName_MapPropertyByString_NotFound()
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
            TestDelegate action = () => importer.Map<SampleClass>(name, "NotExistProperty");

            // Assert
            Assert.Throws<InvalidOperationException>(action);
        }

        [Test]
        public void ColumnName_MapPropertyByString_AmbiguousMatchException()
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
            TestDelegate action = () => importer.Map<TestClass>(name, "myString");

            // Assert
            Assert.Throws<InvalidOperationException>(action);
        }

        [Test]
        public void ColumnsWithSameNameTest()
        {
            // Prepare
            const string str1 = "aBC";
            const string str2 = "aBCd";
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

        [Test]
        public void IgnoredTest()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetSimpleWorkbook(date, str1);

            var header = workbook.GetSheetAt(1).GetRow(0).CreateCell(41);
            header.SetCellValue(nameof(sample.GeneralProperty));
            workbook.GetSheetAt(1).GetRow(1).GetCell(1).SetCellValue(str1);
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);

            var importer = new Mapper(workbook);

            // Act
            importer.Ignore<SampleClass>(o => o.GeneralProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNull(objs[0].Value.GeneralProperty);
            Assert.IsNull(objs[1].Value.GeneralProperty);
        }

        /// <summary>
        /// Test for Issue 1: cannot ignore properties from base class.
        /// </summary>
        [Test]
        public void Issue_1_Test()
        {
            // Prepare
            var sample = new SampleClass();
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook1 = GetSimpleWorkbook(date, str1); // For import.
            var workbook2 = GetBlankWorkbook(); // For export.

            // For flunt method Ignore.
            var header = workbook1.GetSheetAt(1).GetRow(0).CreateCell(41);
            var row = workbook1.GetSheetAt(1).CreateRow(21);
            header.SetCellValue(nameof(sample.BaseStringProperty));
            row.CreateCell(41).SetCellValue(str1);

            // For Ignore Attribute.
            header = workbook1.GetSheetAt(1).GetRow(0).CreateCell(42);
            header.SetCellValue(nameof(sample.BaseIgnoredProperty));
            row.CreateCell(42).SetCellValue(str1);

            var importer = new Mapper(workbook1);
            var exporter = new Mapper(workbook2);

            // Act
            importer.Ignore<SampleClass>(o => o.BaseStringProperty);
            var objs = importer.Take<SampleClass>(1).ToList();

            exporter.Ignore<SampleClass>(o => o.BaseStringProperty);
            sample.BaseStringProperty = "abc";
            exporter.Put(new[] { sample });
            var hasBaseStringProperty = false;
            var hasBaseIgnoredProperty = false;

            foreach (var cell in workbook2.GetSheetAt(0).GetRow(0))
            {
                if (cell.StringCellValue == nameof(sample.BaseStringProperty))
                {
                    hasBaseStringProperty = true;
                    break;
                }
            }

            foreach (var cell in workbook2.GetSheetAt(0).GetRow(0))
            {
                if (cell.StringCellValue == nameof(sample.BaseIgnoredProperty))
                {
                    hasBaseIgnoredProperty = true;
                    break;
                }
            }

            // Assert
            Assert.IsNull(objs[1].Value.BaseStringProperty);
            Assert.IsNull(objs[1].Value.BaseIgnoredProperty);
            Assert.IsFalse(hasBaseStringProperty);
            Assert.IsFalse(hasBaseIgnoredProperty);
        }

        [Test]
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

        [Test]
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

        [Test]
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

        [Test]
        public void Map_IndexAndName_ShouldWork()
        {
            // Arrange
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetSimpleWorkbook(date, str1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(12).SetCellValue("StringProperty");

            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str1);
            workbook.GetSheetAt(1).GetRow(1).CreateCell(12).SetCellValue(str2);

            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(12, "StringProperty");
            var objs = mapper.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(str2, objs[0].Value.StringProperty);
        }

        [Test]
        public void IgnoreErrorsFor_Name_ShouldWork()
        {
            // Arrange
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetSimpleWorkbook(date, str1);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11);

            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str1);

            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(11, "Int32Property");
            mapper.IgnoreErrorsFor<SampleClass>("Int32Property");
            var objs = mapper.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs[0].Value.Int32Property);
            Assert.IsTrue(objs[0].ErrorColumnIndex < 0); // Less than 0 means no error or error ignored.
        }

        [Test]
        public void Ignore_PropertyNames_ShouldIgnored()
        {
            // Arrange
            const string str1 = "aBC";
            const string str2 = "12";
            const string str3 = "EFG";
            var workbook = GetBlankWorkbook();
            var row1 = workbook.GetSheetAt(0).CreateRow(0);
            var row2 = workbook.GetSheetAt(0).CreateRow(1);
            row1.CreateCell(11).SetCellValue("StringProperty");
            row1.CreateCell(12).SetCellValue("Int32Property");
            row1.CreateCell(13).SetCellValue("GeneralProperty");

            row2.CreateCell(11).SetCellValue(str1);
            row2.CreateCell(12).SetCellValue(str2);
            row2.CreateCell(13).SetCellValue(str3);

            var mapper = new Mapper(workbook);

            // Act
            mapper.Ignore<SampleClass>("StringProperty", "GeneralProperty");
            var objs = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNull(objs[0].Value.StringProperty);
            Assert.AreEqual(12, objs[0].Value.Int32Property);
            Assert.IsNull(objs[0].Value.GeneralProperty);
        }

        [Test]
        public void Ignore_DynamicPropertyNames_ShouldIgnored()
        {
            // Arrange
            const string str1 = "aBC";
            const string str2 = "12";
            const string str3 = "EFG";
            var workbook = GetBlankWorkbook();
            var row1 = workbook.GetSheetAt(0).CreateRow(0);
            var row2 = workbook.GetSheetAt(0).CreateRow(1);
            row1.CreateCell(11).SetCellValue("StringProperty");
            row1.CreateCell(12).SetCellValue("Int32Property");
            row1.CreateCell(13).SetCellValue("GeneralProperty");

            row2.CreateCell(11).SetCellValue(str1);
            row2.CreateCell(12).SetCellValue(str2);
            row2.CreateCell(13).SetCellValue(str3);

            var mapper = new Mapper(workbook);

            // Act
            mapper.Ignore<dynamic>(new[] { "StringProperty", "GeneralProperty" });
            var objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNull(objs[0].Value.StringProperty);
            Assert.AreEqual(str2, objs[0].Value.Int32Property);
            Assert.IsNull(objs[0].Value.GeneralProperty);
        }

        private class TestClass
        {
            public string MyString { get; set; }
            public string MYString { get; set; }
        }
    }
}
