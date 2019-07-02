﻿using System;
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

        [TestMethod]
        public void ColumnAttributeIndexTestFromGivenIndexHeader()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            workbook.GetSheetAt(1).GetRow(0).CreateCell(11).SetCellValue("targetColumn");
            workbook.GetSheetAt(1).GetRow(1).CreateCell(11).SetCellValue(str);
            workbook.GetSheetAt(1).ShiftRows(0, 7, 7);
            workbook.GetSheetAt(1).CreateRow(0).CreateCell(0).SetCellValue("Ignored");
            var importer = new Mapper(workbook);
            importer.HeaderRowIndex = 7;
            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            Assert.AreEqual(str, obj.Value.ColumnIndexAttributeProperty);
        }
    }
}
