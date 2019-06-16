using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Npoi.Mapper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using test.Sample;

namespace test
{
    [TestClass]
    public class ImportGeneralTests : TestBase
    {
        [TestInitialize]
        public void InitializeTest()
        {
        }

        [TestCleanup]
        public void CleanupTest()
        {
        }

        private class TestClass
        {
            public string String { get; set; }
            public DateTime DateTime { get; set; }
            public double Double { get; set; }
        }

        private class NullableClass
        {
            public DateTime? NullableDateTime { get; set; }
            public string NormalString { get; set; }
        }

        [TestMethod]
        public void ImporterWithoutAnyMapping()
        {
            // Arrange
            var stream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Mapper(stream);
            var items = importer.Take<TestClass>("TestClass").ToList();

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
            Assert.AreEqual(3, items.Count);
            Assert.IsTrue(items[1].Value.DateTime.Year == 2017);
            Assert.IsTrue(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
        }

        [TestMethod]
        public void ImporterWithFormat()
        {
            // Arrange
            var stream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Mapper(stream);
            importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
            var items = importer.Take<TestClass>("TestClass").ToList();

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
            Assert.AreEqual(3, items.Count);
            Assert.IsTrue(items[1].Value.DateTime.Year == 2017);
            Assert.IsTrue(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
        }

        [TestMethod]
        public void Import_ParseStringToNullableDateTime_Success()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.IsTrue(items[0].Value.NullableDateTime.Value.Year == 2017);
            Assert.IsTrue(items[1].Value.NullableDateTime.Value.Year == 2017);
            Assert.IsTrue(items[2].Value.NullableDateTime.Value.Year == 2017);
        }

        [TestMethod]
        public void Import_ErrorOnNullable_GetNullObject()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.IsNull(items[3].Value);
        }

        [TestMethod]
        public void Import_IgnoreErrorOnNullable_GetNullProperty()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            importer.IgnoreErrorsFor<NullableClass>(o => o.NullableDateTime);
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.IsNull(items[2].Value.NullableDateTime);
            Assert.IsNotNull(items[2].Value.NormalString);
            Assert.IsNull(items[3].Value.NullableDateTime);
            Assert.IsNotNull(items[3].Value.NormalString);
        }

        [TestMethod]
        public void ImporterConstructorWorkbookTest()
        {
            // Arrange
            var workbook = GetSimpleWorkbook(DateTime.MaxValue, "dummy");

            // Act
            var importer = new Mapper(workbook);

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ImporterConstructorNullStreamTest()
        {
            // Arrange
            Stream nullStream = null;

            // Act
            var importer = new Mapper(nullStream);


            // Assert
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ImporterConstructorNullWorkbookTest()
        {
            // Arrange
            IWorkbook nullWorkbook = null;

            // Act
            var importer = new Mapper(nullWorkbook);

            // Assert
        }

        [TestMethod]
        public void ImporterConstructorFilePathTest()
        {
            // Arrange

            // Act
            var importer = new Mapper("Book1.xlsx");


            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public void ImporterConstructorFilePathNotExistTest()
        {
            // Arrange

            // Act
            var importer = new Mapper("dummy.txt");

            // Assert
        }

        [TestMethod]
        public void ImporterNoElementTest()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var header = workbook.CreateSheet("sheet1").CreateRow(0);
            header.CreateCell(0).SetCellValue("StringProperty");
            header.CreateCell(1).SetCellValue("Int32Property");
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(0);

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count());
        }

        [TestMethod]
        public void ImporterEmptySheetTest()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("sheet1");
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(0);

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count());
        }

        [TestMethod]
        public void TakeByHeaderIndexTest()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            var objDate = obj.Value.DateProperty;

            Assert.AreEqual(date.ToLongDateString(), objDate.ToLongDateString());
            Assert.AreEqual(str, obj.Value.StringProperty);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TakeByHeaderIndexOutOfRangeTest()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Mapper(workbook);

            // Act
            // ReSharper disable once UnusedVariable
            var objs = importer.Take<SampleClass>(10).ToList();

            // Assert
        }

        [TestMethod]
        public void TakeByHeaderNameTest()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>("sheet2").ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            var objDate = obj.Value.DateProperty;

            Assert.AreEqual(date.ToLongDateString(), objDate.ToLongDateString());
            Assert.AreEqual(str, obj.Value.StringProperty);
        }

        [TestMethod]
        public void TakeByHeaderNameNotExistTest()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>("notExistSheet").ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count);
        }

        [TestMethod]
        public void Import_ConvertValueError_GetNullObject()
        {
            // Arrange
            const double dou1 = 1.833;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);

            // Header row
            sheet.GetRow(0).CreateCell(0).SetCellValue("DoubleProperty");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Int32Property");
            sheet.GetRow(0).CreateCell(2).SetCellValue("StringProperty");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(dou1);
            sheet.GetRow(1).CreateCell(1).SetCellValue((string)null);
            sheet.GetRow(1).CreateCell(2).SetCellValue(str1);

            // Row #2
            sheet.GetRow(2).CreateCell(0).SetCellValue(dou1);
            sheet.GetRow(2).CreateCell(1).SetCellValue("dummy");
            sheet.GetRow(2).CreateCell(2).SetCellValue(str2);
            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.AreEqual(default(int), items[0].Value.Int32Property);
            Assert.IsNull(items[1].Value);
        }

        [TestMethod]
        public void Import_IgnoreValueTypeParseError_GetDefaultProperty()
        {
            // Arrange
            const double dou1 = 1.833;
            const int int1 = 22;
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);

            // Header row
            sheet.GetRow(0).CreateCell(0).SetCellValue("DoubleProperty");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Int32Property");
            sheet.GetRow(0).CreateCell(2).SetCellValue("StringProperty");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(int1.ToString());
            sheet.GetRow(1).CreateCell(1).SetCellValue(dou1.ToString("f3"));
            sheet.GetRow(1).CreateCell(2).SetCellValue(str1);

            // Row #2
            sheet.GetRow(2).CreateCell(0).SetCellValue("dummy");
            sheet.GetRow(2).CreateCell(1).SetCellValue("dummy");
            sheet.GetRow(2).CreateCell(2).SetCellValue(str2);
            var mapper = new Mapper(workbook);

            // Act
            mapper.IgnoreErrorsFor<SampleClass>(o => o.DoubleProperty);
            mapper.IgnoreErrorsFor<SampleClass>(o => o.Int32Property);
            var items = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.AreEqual(int1, items[0].Value.DoubleProperty);
            Assert.AreEqual(Math.Round(dou1), items[0].Value.Int32Property);
            Assert.AreEqual(str1, items[0].Value.StringProperty);
            Assert.AreEqual(default(double), items[1].Value.DoubleProperty);
            Assert.AreEqual(default(int), items[1].Value.Int32Property);
            Assert.AreEqual(str2, items[1].Value.StringProperty);
        }

        [TestMethod]
        public void Import_ValidEnum_ShouldWork()
        {
            // Arrange
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);
            sheet.CreateRow(3);

            // Header row
            sheet.GetRow(0).CreateCell(0).SetCellValue("EnumProperty");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(SampleEnum.Value1.ToString());

            // Row #2
            sheet.GetRow(2).CreateCell(0).SetCellValue(SampleEnum.Value2.ToString());

            // Row #3
            sheet.GetRow(3).CreateCell(0).SetCellValue("value3");

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<SampleClass>().ToList();

            // Assert
            Assert.AreEqual(SampleEnum.Value1, items[0].Value.EnumProperty);
            Assert.AreEqual(SampleEnum.Value2, items[1].Value.EnumProperty);
            Assert.AreEqual(SampleEnum.Value3, items[2].Value.EnumProperty);
        }

        [TestMethod]
        public void Map_WithIndexAndName_ShouldImportByIndex()
        {
            // Arrange
            var workbook = GetEmptyWorkbook();
            const string nameString = "StringProperty";
            const string nameGeneral = "GeneralProperty";
            var sheet = workbook.CreateSheet();

            var headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue(nameGeneral);
            headerRow.CreateCell(1).SetCellValue(nameString);

            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("a");
            row1.CreateCell(1).SetCellValue("b");

            var mapper = new Mapper(workbook);

            // Act
            mapper.Map<SampleClass>(0, "StringProperty", nameString);
            mapper.Map<SampleClass>(1, "GeneralProperty", nameGeneral);
            var obj = mapper.Take<SampleClass>().Select(o=>o.Value).ToArray()[0];

            // Assert
            Assert.AreEqual("a", obj.StringProperty);
            Assert.AreEqual("b", obj.GeneralProperty);
        }

        [TestMethod]
        public void ImporterWithoutAnyMappingFromGivenHeaderIndex()
        {
            // Arrange
            CreateShiftedRowsWorkbook("Book1.xlsx", "Book6.xlsx", "TestClass", 4);
            var stream = new FileStream("Book6.xlsx", FileMode.Open);


            // Act
            var importer = new Mapper(stream) {HeaderRowIndex =  4};
            var items = importer.Take<TestClass>("TestClass").ToList();

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
            Assert.AreEqual(3, items.Count);
            Assert.IsTrue(items[1].Value.DateTime.Year == 2017);
            Assert.IsTrue(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
        }

        [TestMethod]
        public void ImporterNoElementTestFromGivenHeaderIndex()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var header = workbook.CreateSheet("sheet1").CreateRow(0);
            header.CreateCell(0).SetCellValue("StringProperty");
            header.CreateCell(1).SetCellValue("Int32Property");
            const int importerHeaderRowIndex = 9;
            workbook.GetSheet("sheet1").ShiftRows(0, importerHeaderRowIndex, importerHeaderRowIndex);
            var importer = new Mapper(workbook);
            importer.HeaderRowIndex = importerHeaderRowIndex;

            // Act
            var objs = importer.Take<SampleClass>(0);

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count());
        }
    }
}
