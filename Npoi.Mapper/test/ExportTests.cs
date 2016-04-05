using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;
using test.Sample;

namespace test
{
    [TestClass]
    public class ExportTests : TestBase
    {
        SampleClass sampleObj = new SampleClass
        {
            BuiltinFormatProperty = DateTime.Today,
            ColumnIndexAttributeProperty = "Column Index",
            CustomFormatProperty = 0.87,
            DateProperty = DateTime.Now,
            DoubleProperty = 78,
            GeneralProperty = "general sting",
            BoolProperty = true,
            EnumProperty = SampleEnum.Value3,
            IgnoredAttributeProperty = "Ignored column",
            Int32Property = 100,
            SingleColumnResolverProperty = "I'm here..."
        };

        [TestMethod]
        public void SaveSheetTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save<SampleClass>("test.xlsx", 1);

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }

        [TestMethod]
        public void SaveObjectsTest()
        {
            // Prepare
            var exporter = new Mapper();
            exporter.Map<SampleClass>("General Column", o => o.GeneralProperty);

            // Act
            exporter.Save("test.xlsx", new[] { sampleObj }, "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
        }

        [TestMethod]
        public void SaveTrackedObjectsTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
        }

        [TestMethod]
        public void FormatAttributeTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.BuiltinFormatProperty = DateTime.Now;
            objs[0].Value.CustomFormatProperty = 100.234;

            // Act
            exporter.Map<SampleClass>(11, o => o.BuiltinFormatProperty);
            exporter.Map<SampleClass>(12, o => o.CustomFormatProperty);
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            var dateStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(11).CellStyle;
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(0xf, dateStyle.DataFormat);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);
        }

        [TestMethod]
        public void FormatMethodTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.DateProperty = DateTime.Now;
            objs[0].Value.DoubleProperty = 100.234;

            // Act
            exporter.Map<SampleClass>(11, o => o.DateProperty);
            exporter.Map<SampleClass>(12, o => o.DoubleProperty);
            exporter.Format<SampleClass>(0xf, o => o.DateProperty);
            exporter.Format<SampleClass>("0%", o => o.DoubleProperty);
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            var dateStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(11).CellStyle;
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(0xf, dateStyle.DataFormat);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);
        }

        [TestMethod]
        public void NoHeaderTest()
        {
            // Prepare
            var exporter = new Mapper { HasHeader = false };
            const string sheetName = "newSheet";

            // Act
            exporter.Save("test.xlsx", new[] { sampleObj, }, sheetName);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(1, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
        }

        [TestMethod]
        public void ExportToExistingFileTest()
        {
            // Prepare
            var exporter = new Mapper();
            const string fileName = "test.xlsx";
            const string sheetName = "oldSheet";
            exporter.Save(fileName, new[] { sampleObj, }, sheetName);
            exporter.Workbook.CreateSheet("newSheet");

            // Act
            exporter.Save(fileName, new[] { sampleObj, }, sheetName, true, false);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.NumberOfSheets);
        }

        [TestMethod]
        public void ExportToNewFileTest()
        {
            // Prepare
            var exporter = new Mapper();
            const string fileName = "test.xlsx";
            const string sheetName = "oldSheet";
            exporter.Save(fileName, new[] { sampleObj, }, sheetName);
            exporter.Workbook.CreateSheet("newSheet");

            // Act
            exporter.Save(fileName, new[] { sampleObj, }, sheetName, true, true);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(1, exporter.Workbook.NumberOfSheets);
        }
    }
}
