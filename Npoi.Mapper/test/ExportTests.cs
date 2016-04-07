using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
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

        const string FileName = "test.xlsx";

        [TestMethod]
        public void SaveSheetTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save<SampleClass>(FileName, 1);

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);

            // Cleanup
            File.Delete(FileName);
        }

        [TestMethod]
        public void SaveObjectsTest()
        {
            // Prepare
            var exporter = new Mapper();
            exporter.Map<SampleClass>("General Column", o => o.GeneralProperty);

            // Act
            exporter.Save(FileName, new[] { sampleObj }, "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);

            // Cleanup
            File.Delete(FileName);
        }

        [TestMethod]
        public void SaveTrackedObjectsTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);

            // Cleanup
            File.Delete(FileName);
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
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            var dateStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(11).CellStyle;
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(0xf, dateStyle.DataFormat);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);

            // Cleanup
            File.Delete(FileName);
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
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            var dateStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(11).CellStyle;
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(0xf, dateStyle.DataFormat);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);

            // Cleanup
            File.Delete(FileName);
        }

        [TestMethod]
        public void NoHeaderTest()
        {
            // Prepare
            var exporter = new Mapper { HasHeader = false };
            const string sheetName = "newSheet";

            // Act
            exporter.Save(FileName, new[] { sampleObj, }, sheetName);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(1, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);

            // Cleanup
            File.Delete(FileName);
        }

        [TestMethod]
        public void ExportXlsTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "newSheet";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();

            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, true, false);

            // Assert
            Assert.IsNotNull(exporter.Workbook as HSSFWorkbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);

            // Cleanup
            File.Delete(existingFile);
        }

        [TestMethod]
        public void OverwriteNewFileTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();

            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, true);

            // Assert
            Assert.AreEqual(1, exporter.Workbook.NumberOfSheets);

            // Cleanup
            File.Delete(existingFile);
        }

        [TestMethod]
        public void MergeToExistedRowsTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();
            exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
            exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, false);

            // Assert
            var sheet = exporter.Workbook.GetSheet(sheetName);
            Assert.AreEqual(sampleObj.GeneralProperty, sheet.GetRow(4).GetCell(1).StringCellValue);
            Assert.AreEqual(sampleObj.DateProperty.Date, sheet.GetRow(4).GetCell(2).DateCellValue.Date);

            // Cleanup
            File.Delete(existingFile);
        }

        [TestMethod]
        public void PutAppendRowTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper(existingFile);
            exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
            exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

            // Act
            exporter.Put(new[] { sampleObj, }, sheetName, false);

            // Assert
            var sheet = exporter.Workbook.GetSheet(sheetName);
            Assert.AreEqual(sampleObj.GeneralProperty, sheet.GetRow(4).GetCell(1).StringCellValue);
            Assert.AreEqual(sampleObj.DateProperty.Date, sheet.GetRow(4).GetCell(2).DateCellValue.Date);

            // Cleanup
            File.Delete(existingFile);
        }

        [TestMethod]
        public void PutOverwriteRowTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "Allocations";
            if(File.Exists(existingFile))File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper(existingFile);
            exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
            exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

            // Act
            exporter.Put(new[] { sampleObj, }, sheetName, true);

            // Assert
            var sheet = exporter.Workbook.GetSheet(sheetName);
            Assert.AreEqual(sampleObj.GeneralProperty, sheet.GetRow(1).GetCell(1).StringCellValue);
            Assert.AreEqual(sampleObj.DateProperty.Date, sheet.GetRow(1).GetCell(2).DateCellValue.Date);

            // Cleanup
            File.Delete(existingFile);
        }

        [TestMethod]
        public void SaveWorkbookToFileTest()
        {
            // Prepare
            const string existingFile = "Book2.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            
            var exporter = new Mapper("Book1.xlsx");

            // Act
            exporter.Save(existingFile);

            // Assert
            Assert.IsTrue(File.Exists(existingFile));

            // Cleanup
            File.Delete(existingFile);

        }
    }
}
