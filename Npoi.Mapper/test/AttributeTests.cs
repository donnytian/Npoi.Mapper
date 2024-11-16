using System;
using System.Linq;
using Npoi.Mapper;
using NUnit.Framework;
using test.Sample;

namespace test
{
    [TestFixture]
    public class AttributeTests : TestBase
    {
        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs, Has.Count.EqualTo(1));

            var obj = objs[0];
            Assert.That(obj.Value.ColumnIndexAttributeProperty, Is.EqualTo(str));
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count, Is.EqualTo(1));

            var obj = objs[0];
            Assert.That(obj.Value.ColumnNameAttributeProperty, Is.EqualTo(str));
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs, Has.Count.EqualTo(1));

            var obj = objs[0];
            Assert.That(obj.Value.DisplayNameProperty, Is.EqualTo(str));
        }

        [Test]
        public void UseLastNonBlankValueAttributeTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetSimpleWorkbook(date, str1);

            var header = workbook.GetSheetAt(1).GetRow(0).CreateCell(41);
            header.SetCellValue(nameof(SampleClass.UseLastNonBlankValueAttributeProperty));

            // Create 4 rows, row 22 and 23 have empty values.
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);
            workbook.GetSheetAt(1).CreateRow(22).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(23).CreateCell(41).SetCellValue(string.Empty);
            workbook.GetSheetAt(1).CreateRow(24).CreateCell(41).SetCellValue(str2);

            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count, Is.EqualTo(5));

            var obj = objs[1];
            Assert.That(obj.Value.UseLastNonBlankValueAttributeProperty, Is.EqualTo(str1));

            obj = objs[2];
            Assert.That(obj.Value.UseLastNonBlankValueAttributeProperty, Is.EqualTo(str1));

            obj = objs[3];
            Assert.That(obj.Value.UseLastNonBlankValueAttributeProperty, Is.EqualTo(str1));

            obj = objs[4];
            Assert.That(obj.Value.UseLastNonBlankValueAttributeProperty, Is.EqualTo(str2));
        }

        [Test]
        public void IgnoreAttributeTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetSimpleWorkbook(date, str1);

            workbook.GetSheetAt(1).GetRow(0).CreateCell(41).SetCellValue(nameof(SampleClass.IgnoredAttributeProperty));
            workbook.GetSheetAt(1).CreateRow(21).CreateCell(41).SetCellValue(str1);

            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.That(objs[0].Value.IgnoredAttributeProperty, Is.Null);
        }
    }
}
