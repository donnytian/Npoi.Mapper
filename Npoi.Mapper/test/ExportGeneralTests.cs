using System;
using System.IO;
using System.Linq;
using Npoi.Mapper;
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
            exporter.Save(FileName, new[] { dummyObj }, sheetName, false);
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);

            // Assert
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows, Is.EqualTo(2));
            Assert.That(DateUtil.IsCellDateFormatted(dateCell));
            Assert.That(exporter.Take<DummyClass>(sheetName).First().Value.String, Is.EqualTo(dummyObj.String));
            Assert.That(exporter.Take<DummyClass>(sheetName).First().Value.Double, Is.EqualTo(dummyObj.Double));
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
            exporter.Save(FileName, new[] { dummyObj }, sheetName, false);
            var items = exporter.Take<DummyClass>(sheetName).ToList();
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);

            // Assert
            Assert.That(exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows, Is.EqualTo(2));
            Assert.That(DateUtil.IsCellDateFormatted(dateCell));
            Assert.That(items.First().Value.DateTime.ToLongDateString(), Is.EqualTo(dummyObj.DateTime.ToLongDateString()));
            Assert.That(items.First().Value.Double, Is.EqualTo(dummyObj.Double));
            Assert.That(items.First().Value.DateTime2.ToLongDateString(), Is.EqualTo(dummyObj.DateTime2.ToLongDateString()));
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
            exporter.Save(FileName, new[] { obj1, obj2 }, sheetName, true);

            var items = exporter.Take<NullableClass>(sheetName).ToList();
            var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(2).GetCell(0);

            // Assert
            Assert.That(exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows, Is.EqualTo(3));
            Assert.That(items.First().Value.DummyString, Is.EqualTo(obj1.DummyString));
            Assert.That(items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString(), Is.EqualTo(obj2.NullableDateTime.Value.ToLongDateString()));
            Assert.That(DateUtil.IsCellDateFormatted(dateCell));
            Assert.That(items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString(), Is.EqualTo(obj2.NullableDateTime.Value.ToLongDateString()));
            Assert.That(exporter.Take<NullableClass>(sheetName).First().Value.NullableDateTime.HasValue, Is.False);
        }

        [Test]
        public void SaveObjectsTest()
        {
            // Prepare
            var exporter = new Mapper();
            exporter.Map<SampleClass>("General Column", o => o.GeneralProperty);
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.Save(FileName, new[] { sampleObj }, "newSheet", false);

            // Assert
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows, Is.EqualTo(2));
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
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet", false);

            // Assert
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows, Is.EqualTo(2));
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
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet", false);

            // Assert
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(doubleStyle.DataFormat, Is.Not.EqualTo(0));
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
            exporter.Save(FileName, objs.Select(info => info.Value), "newSheet", false);

            // Assert
            var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(doubleStyle.DataFormat, Is.Not.EqualTo(0));
        }

        [Test]
        public void NoHeaderTest()
        {
            // Prepare
            var exporter = new Mapper { HasHeader = false };
            const string sheetName = "newSheet";
            if (File.Exists(FileName)) File.Delete(FileName);

            // Act
            exporter.Save(FileName, new[] { sampleObj, }, sheetName, false);

            // Assert
            Assert.That(exporter.Workbook, Is.Not.Null);
            Assert.That(exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows, Is.EqualTo(1));
        }

        [Test]
        public void ExportXlsxTest()
        {
            // Prepare
            const string existingFile = "ExportXlsxTest.xlsx";
            const string sheetName = "newSheet";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();

            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, false, false);

            // Assert
            Assert.That(exporter.Workbook as XSSFWorkbook, Is.Not.Null);
            Assert.That(exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows, Is.EqualTo(2));
            File.Delete(existingFile);
        }

        [Test]
        public void OverwriteRowsInExistingFileTest()
        {
            // Prepare
            const string existingFile = "OverwriteRowsInExistingFileTest.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();
            exporter.Map<SampleClass>("Resource Name", c => c.Int32Property);
            
            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, false, overwrite: true);

            // Assert
            var sheet = exporter.Workbook.GetSheet(sheetName);
            var cellValue = sheet.GetRow(1).Cells[0].NumericCellValue;
            Assert.That(exporter.Workbook.NumberOfSheets > 1, Is.True);
            Assert.That(sheet.LastRowNum == 1, Is.True);
            Assert.That(cellValue, Is.EqualTo(sampleObj.Int32Property));
            File.Delete(existingFile);
        }

        [Test]
        public void MergeToExistedRowsTest()
        {
            // Prepare
            const string existingFile = "MergeToExistedRowsTest.xlsx";
            const string sheetName = "Allocations";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy("Book1.xlsx", existingFile);
            var exporter = new Mapper();
            exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
            exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

            // Act
            exporter.Save(existingFile, new[] { sampleObj, }, sheetName, false, overwrite: false);

            // Assert
            var sheet = exporter.Workbook.GetSheet(sheetName);
            Assert.That(sheet.GetRow(4).GetCell(1).StringCellValue, Is.EqualTo(sampleObj.GeneralProperty));
            Assert.That(sheet.GetRow(4).GetCell(2).DateCellValue?.Date, Is.EqualTo(sampleObj.DateProperty.Date));
            File.Delete(existingFile);
        }

        [Test]
        public void PutAppendRowTest()
        {
            // Prepare
            const string existingFile = "PutAppendRowTest.xlsx";
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
            Assert.That(sheet.GetRow(4).GetCell(1).StringCellValue, Is.EqualTo(sampleObj.GeneralProperty));
            Assert.That(sheet.GetRow(4).GetCell(2).DateCellValue?.Date, Is.EqualTo(sampleObj.DateProperty.Date));
            File.Delete(existingFile);
        }

        [Test]
        public void PutOverwriteRowTest()
        {
            // Prepare
            const string existingFile = "PutOverwriteRowTest.xlsx";
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
            var row = sheet.GetRow(1);
            Assert.That(sheet.LastRowNum, Is.EqualTo(1));
            Assert.That(row.GetCell(1).StringCellValue, Is.EqualTo(sampleObj.GeneralProperty));
            Assert.That(row.GetCell(2).DateCellValue?.Date, Is.EqualTo(sampleObj.DateProperty.Date));
        }

        [Test]
        public void SaveWorkbookToFileTest()
        {
            // Prepare
            const string fileName = "temp4.xlsx";
            if (File.Exists(fileName)) File.Delete(fileName);

            var exporter = new Mapper("Book1.xlsx");

            // Act
            exporter.Save(fileName, false);

            // Assert
            Assert.That(File.Exists(fileName), Is.True);
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
            Assert.That(workbook.NumberOfSheets > 0, Is.True);
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
            Assert.That(workbook.NumberOfSheets > 0, Is.True);
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
            Assert.That(row.GetCell(0).StringCellValue, Is.EqualTo(nameString));
            Assert.That(row.GetCell(1).StringCellValue, Is.EqualTo(nameInt));
            Assert.That(row.GetCell(2).StringCellValue, Is.EqualTo(nameBool));
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
            Assert.That(sheet.PhysicalNumberOfRows, Is.EqualTo(1 + (hasHeader ? 1 : 0)));
            Assert.That(row.GetCell(0).StringCellValue, Is.EqualTo(nameString));
        }

        [Test]
        public void TakeZeroRow_Then_PutZeroObject_VerifyHeaders()
        {
            // Arrange
            var workbook = GetEmptyWorkbook();
            var sheet = workbook.CreateSheet();
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(SampleClass.BoolProperty));
            header.CreateCell(1).SetCellValue(nameof(SampleClass.StringProperty));

            var mapper = new Mapper(workbook);

            // Act
            var objs = mapper.Take<SampleClass>().Select(x => x.Value);
            mapper.Put(objs, "Sheet1");

            // Assert
            var result = WriteAndReadBack(mapper.Workbook);
            var row = result.GetSheetAt(0).GetRow(0);
            Assert.That(row.Cells.Count > 0);
            Assert.That(string.IsNullOrWhiteSpace(row.Cells[0].StringCellValue), Is.False);
        }
    }
}
