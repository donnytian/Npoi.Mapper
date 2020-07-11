using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using test.Sample;

namespace test
{
    [TestFixture]
    public class ExportGeneralTests : TestBase
    {
        SampleClass sampleObj = new SampleClass
        {
            ColumnIndexAttributeProperty = "Column Index",
            CustomFormatProperty = 0.87,
            DateProperty = DateTime.Now,
            DoubleProperty = 78,
            GeneralProperty = "general sting",
            StringProperty = "balabala",
            BoolProperty = true,
            EnumProperty = SampleEnum.Value3,
            IgnoredAttributeProperty = "Ignored column",
            Int32Property = 100,
            SingleColumnResolverProperty = "I'm here..."
        };

        private class DummyClass
        {
            public string String { get; set; }
            public DateTime DateTime { get; set; }
            public double Double { get; set; }
            public DateTime DateTime2 { get; set; }
        }

        private DummyClass dummyObj = new DummyClass
        {
            String = "My string",
            DateTime = DateTime.Now,
            Double = 0.4455,
            DateTime2 = DateTime.Now.AddDays(1)
        };

        private class NullableClass
        {
            public DateTime? NullableDateTime { get; set; }
            public string DummyString { get; set; }
        }

        const string FileName = "test.xlsx";

        [Test]
        public void SaveSheetWithoutAnyMapping()
        {
            // Arrange
            var exporter = new Mapper();
            var sheetName = "newSheet";
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.Save(FileName, new[] { dummyObj }, sheetName);
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
            Assert.IsTrue(DateUtil.IsCellDateFormatted(dateCell));
            Assert.AreEqual(dummyObj.String, exporter.Take<DummyClass>(sheetName).First().Value.String);
            Assert.AreEqual(dummyObj.Double, exporter.Take<DummyClass>(sheetName).First().Value.Double);
        }

        [Test]
        public void SaveSheetUseFormat()
        {
            // Arrange
            var exporter = new Mapper();
            var sheetName = "newSheet";
            var dateFormat = "yyyy.MM.dd hh.mm.ss";
            var doubleFormat = "0%";
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.UseFormat(typeof(DateTime), dateFormat);
            exporter.UseFormat(typeof(double), doubleFormat);
            exporter.Save(FileName, new[] { dummyObj }, sheetName);
            var items = exporter.Take<DummyClass>(sheetName).ToList();
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);

            // Assert
            Assert.AreEqual(2, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
            Assert.IsTrue(DateUtil.IsCellDateFormatted(dateCell));
            Assert.AreEqual(dummyObj.DateTime.ToLongDateString(), items.First().Value.DateTime.ToLongDateString());
            Assert.AreEqual(dummyObj.Double, items.First().Value.Double);
            Assert.AreEqual(dummyObj.DateTime2.ToLongDateString(), items.First().Value.DateTime2.ToLongDateString());
        }

        [Test]
        public void SaveSheetUseFormatForNullable()
        {
            // Arrange
            var exporter = new Mapper();
            var sheetName = "newSheet";
            var dateFormat = "yyyy.MM.dd hh.mm.ss";
            var obj1 = new NullableClass { NullableDateTime = null, DummyString = "dummy" };
            var obj2 = new NullableClass { NullableDateTime = DateTime.Now };
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.UseFormat(typeof(DateTime?), dateFormat);

            // Issue #5, if the first data row has null value, then next rows will not be formated
            // So here we make the first date row has a null value for DateTime? property.
            exporter.Save(FileName, new[] { obj1, obj2 }, sheetName);

            var items = exporter.Take<NullableClass>(sheetName).ToList();
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(2).GetCell(0);

            // Assert
            Assert.AreEqual(3, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
            Assert.AreEqual(obj1.DummyString, items.First().Value.DummyString);
            Assert.AreEqual(obj2.NullableDateTime.Value.ToLongDateString(), items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString());
            Assert.IsTrue(DateUtil.IsCellDateFormatted(dateCell));
            Assert.AreEqual(obj2.NullableDateTime.Value.ToLongDateString(), items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString());
            Assert.IsFalse(exporter.Take<NullableClass>(sheetName).First().Value.NullableDateTime.HasValue);
        }

        [Test]
        public void SaveSheetTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            if (File.Exists(FileName)) File.Delete(FileName);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save<SampleClass>(FileName, 1);

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }

        [Test]
        public void SaveObjectsTest()
        {
            // Prepare
            var exporter = new Mapper();
            exporter.Map<SampleClass>("General Column", o => o.GeneralProperty);
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.Save(FileName, new[] { sampleObj }, "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
        }

        [Test]
        public void SaveTrackedObjectsTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            if (File.Exists(FileName)) File.Delete(FileName);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
        }

        [Test]
        public void FormatAttributeTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            if (File.Exists(FileName)) File.Delete(FileName);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.CustomFormatProperty = 100.234;

            // Act
            exporter.Map<SampleClass>(12, o => o.CustomFormatProperty);
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);
        }

        [Test]
        public void FormatMethodTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            if (File.Exists(FileName)) File.Delete(FileName);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.DoubleProperty = 100.234;

            // Act
            exporter.Map<SampleClass>(11, o => o.DateProperty);
            exporter.Map<SampleClass>(12, o => o.DoubleProperty);
            exporter.Format<SampleClass>("0%", o => o.DoubleProperty);
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

            // Assert
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreNotEqual(0, doubleStyle.DataFormat);
        }

        [Test]
        public void NoHeaderTest()
        {
            // Prepare
            var exporter = new Mapper { HasHeader = false };
            const string sheetName = "newSheet";
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.Save(FileName, new[] { sampleObj, }, sheetName);

            // Assert
            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(1, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
        }

        [Test]
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
        }

        [Test]
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
        }

        [Test]
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
        }

        [Test]
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
            var workbook = WriteAndReadBack(exporter.Workbook, existingFile);

            // Assert
            var sheet = workbook.GetSheet(sheetName);
            Assert.AreEqual(sampleObj.GeneralProperty, sheet.GetRow(4).GetCell(1).StringCellValue);
            Assert.AreEqual(sampleObj.DateProperty.Date, sheet.GetRow(4).GetCell(2).DateCellValue.Date);
        }

        [Test]
        public void PutOverwriteRowTest()
        {
            // Prepare
            const string existingFile = "Book3.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper(existingFile);
            exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
            exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);
            exporter.Map<SampleClass>("Name", o => o.StringProperty);
            exporter.Map<SampleClass>("email", o => o.BoolProperty);

            // Act
            exporter.Put(new[] { sampleObj, }, sheetName, true);
            exporter.Put(new[] { sampleObj }, "Resources");
            var workbook = WriteAndReadBack(exporter.Workbook, existingFile);

            // Assert
            var sheet = workbook.GetSheet(sheetName);
            Assert.AreEqual(sampleObj.GeneralProperty, sheet.GetRow(1).GetCell(1).StringCellValue);
            Assert.AreEqual(sampleObj.DateProperty.Date, sheet.GetRow(1).GetCell(2).DateCellValue.Date);
        }

        [Test]
        public void SaveWorkbookToFileTest()
        {
            // Prepare
            const string fileName = "temp4.xlsx";
            if (File.Exists(fileName)) File.Delete(fileName);

            var exporter = new Mapper("Book1.xlsx");

            // Act
            exporter.Save(fileName);

            // Assert
            Assert.IsTrue(File.Exists(fileName));
            File.Delete(fileName);
        }

        // https://github.com/donnytian/Npoi.Mapper/issues/16
        [Test]
        public void PutWithNotExistedSheetIndex_ShouldAutoPopulateSheets()
        {
            // Arrange
            var workbook = GetEmptyWorkbook();

            var mapper = new Mapper(workbook);

            // Act
            mapper.Put(new[] { new object(), }, 100);

            // Assert
            Assert.IsTrue(workbook.NumberOfSheets > 0);
        }

        [Test]
        public void PutWithNotExistedSheetName_ShouldAutoPopulateSheets()
        {
            // Arrange
            var workbook = GetEmptyWorkbook();

            var mapper = new Mapper(workbook);

            // Act
            mapper.Put(new[] { new object(), }, "sheet100");

            // Assert
            Assert.IsTrue(workbook.NumberOfSheets > 0);
        }

        [Test]
        public void Map_WithIndexAndName_ShouldExportCustomColumnName()
        {
            // Arrange
            var workbook = GetEmptyWorkbook();
            const string nameString = "string";
            const string nameInt = "int";
            const string nameBool = "bool";
            var sheet = workbook.CreateSheet();

            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(0, o => o.StringProperty, nameString);
            mapper.Map<SampleClass>(1, o => o.Int32Property, nameInt);
            mapper.Map<SampleClass>(2, o => o.BoolProperty, nameBool);
            mapper.Put(new[] { new SampleClass(), }, 0);

            // Assert
            var row = sheet.GetRow(0);
            Assert.AreEqual(nameString, row.GetCell(0).StringCellValue);
            Assert.AreEqual(nameInt, row.GetCell(1).StringCellValue);
            Assert.AreEqual(nameBool, row.GetCell(2).StringCellValue);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void Put_WithFirstRowIndex_ShouldExportExpectedRows(bool hasHeader)
        {
            // Arrange
            const int firstRowIndex = 100;
            const string nameString = "StringProperty";
            var workbook = GetEmptyWorkbook();
            var sheet = workbook.CreateSheet();

            var item = new SampleClass { StringProperty = nameString };
            var mapper = new Mapper(workbook) { HasHeader = hasHeader, FirstRowIndex = firstRowIndex };
            mapper.Map<SampleClass>(0, o => o.StringProperty, "a");

            // Act
            mapper.Put(new[] { item }, 0);

            // Assert
            var firstDataRowIndex = hasHeader ? firstRowIndex + 1 : firstRowIndex;
            var row = sheet.GetRow(firstDataRowIndex);
            Assert.AreEqual(1 + (hasHeader ? 1 : 0), sheet.PhysicalNumberOfRows);
            Assert.AreEqual(nameString, row.GetCell(0).StringCellValue);
        }
    }
}
