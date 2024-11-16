using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using test.Sample;

namespace test
{
    [TestFixture]
    public class ImportGeneralTests : TestBase
    {
        private class TestClass
        {
            public string String { get; set; }
            public DateTime DateTime { get; set; }
            public double Double { get; set; }
            public DateTimeOffset DateTimeOffsetProperty { get; set; }
        }

        private class NullableClass
        {
            public DateTime? NullableDateTime { get; set; }
            public string NormalString { get; set; }
            public DateTimeOffset? NullableDateTimeOffset { get; set; }
        }

        private class TestDefaultClass
        {
            public string Name { get; set; }

            [DefaultValue(true)]
            public bool AllowEmails { get; set; }
            
            [Column(DefaultValue = true)]
            public bool UseDefaultEmail { get; set; }
            
            [DefaultValue(1)]
            public double HouseHoldNumber { get; set; }
            
            [DefaultValue("P")]
            public string Type { get; set; }

        }

        [Test]
        public void ImporterWithoutAnyMapping()
        {
            // Arrange
            var stream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Mapper(stream);
            var items = importer.Take<TestClass>("TestClass").ToList();

            // Assert
            Assert.That(importer, Is.Not.Null);
            Assert.That(importer.Workbook, Is.Not.Null);
            Assert.That(items.Count, Is.EqualTo(3));
            Assert.That(items[1].Value.DateTime.Year == 2017);
            Assert.That(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001, Is.True);
        }

        [Test]
        public void ImporterWithDefaultValue()
        {
            // Arrange
            using (var stream = new FileStream("test_default.xlsx", FileMode.Open))
            {
                // Act
                var importer = new Mapper(stream) { UseDefaultValueAttribute = true };
                var items = importer.Take<TestDefaultClass>("Sheet1").ToList();

                // Assert
                Assert.That(importer, Is.Not.Null);
                Assert.That(importer.Workbook, Is.Not.Null);
                Assert.That(items.Count, Is.EqualTo(3));
                Assert.That(items[0].Value.AllowEmails, Is.True);
                Assert.That(items[0].Value.UseDefaultEmail, Is.False);
                Assert.That(items[0].Value.HouseHoldNumber, Is.EqualTo(1));
            }
        }

        [Test]
        public void ImporterWithFormat()
        {
            // Arrange
            var stream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Mapper(stream);
            importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
            var items = importer.Take<TestClass>("TestClass").ToList();

            // Assert
            Assert.That(importer, Is.Not.Null);
            Assert.That(importer.Workbook, Is.Not.Null);
            Assert.That(items.Count, Is.EqualTo(3));
            Assert.That(items[1].Value.DateTime.Year == 2017);
            Assert.That(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
        }

        [Test]
        public void Import_ParseStringToNullableDateTime_Success()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.That(items[0].Value.NullableDateTime.Value.Year == 2017);
            Assert.That(items[1].Value.NullableDateTime.Value.Year == 2017);
            Assert.That(items[2].Value.NullableDateTime.Value.Year == 2017);
        }

        [Test]
        public void Import_ErrorOnNullable_GetNullObject()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.That(items[3].ErrorColumnIndex, Is.EqualTo(0));
            Assert.That(items[3].Value.NullableDateTime, Is.Null);
        }

        [Test]
        public void Import_IgnoreErrorOnNullable_GetNullProperty()
        {
            // Arrange
            var importer = new Mapper("Book1.xlsx");

            // Act
            importer.IgnoreErrorsFor<NullableClass>(o => o.NullableDateTime);
            var items = importer.Take<NullableClass>("NullableClass").ToList();

            // Assert
            Assert.That(items[2].Value.NullableDateTime, Is.Null);
            Assert.That(items[2].Value.NormalString, Is.Not.Null);
            Assert.That(items[3].Value.NullableDateTime, Is.Null);
            Assert.That(items[3].Value.NormalString, Is.Not.Null);
        }

        [Test]
        public void ImporterConstructorWorkbookTest()
        {
            // Arrange
            var workbook = GetSimpleWorkbook(DateTime.MaxValue, "dummy");

            // Act
            var importer = new Mapper(workbook);

            // Assert
            Assert.That(importer, Is.Not.Null);
            Assert.That(importer.Workbook, Is.Not.Null);
        }

        [Test]
        public void ImporterConstructorNullStreamTest()
        {
            // Arrange
            Stream nullStream = null;

            // Act
            TestDelegate action = () => new Mapper(nullStream);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void ImporterConstructorNullWorkbookTest()
        {
            // Arrange
            IWorkbook nullWorkbook = null;

            // Act
            TestDelegate action = () => new Mapper(nullWorkbook);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void ImporterConstructorFilePathTest()
        {
            // Arrange

            // Act
            var importer = new Mapper("Book1.xlsx");


            // Assert
            Assert.That(importer, Is.Not.Null);
            Assert.That(importer.Workbook, Is.Not.Null);
        }

        [Test]
        public void ImporterConstructorFilePathNotExistTest()
        {
            // Arrange

            // Act
            TestDelegate action = () => new Mapper("dummy.txt");

            // Assert
            Assert.Throws<FileNotFoundException>(action);
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count(), Is.EqualTo(0));
        }

        [Test]
        public void ImporterEmptySheetTest()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("sheet1");
            var importer = new Mapper(workbook);

            // Act
            var objs = importer.Take<SampleClass>(0);

            // Assert
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count(), Is.EqualTo(0));
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count, Is.EqualTo(1));

            var obj = objs[0];
            var objDate = obj.Value.DateProperty;

            Assert.That(objDate.ToLongDateString(), Is.EqualTo(date.ToLongDateString()));
            Assert.That(obj.Value.StringProperty, Is.EqualTo(str));
        }

        [Test]
        public void TakeByHeaderIndexOutOfRangeTest()
        {
            // Arrange
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Mapper(workbook);

            // Act
            TestDelegate action = () => importer.Take<SampleClass>(10);

            // Assert
            Assert.Throws<ArgumentException>(action);
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count, Is.EqualTo(1));

            var obj = objs[0];
            var objDate = obj.Value.DateProperty;

            Assert.That(objDate.ToLongDateString(), Is.EqualTo(date.ToLongDateString()));
            Assert.That(obj.Value.StringProperty, Is.EqualTo(str));
        }

        [Test]
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
            Assert.That(objs, Is.Not.Null);
            Assert.That(objs.Count, Is.EqualTo(0));
        }

        [Test]
        public void Import_ConvertValueError_GotErrorColumnIndex()
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
            Assert.That(items[0].Value.Int32Property, Is.EqualTo(default(int)));
            Assert.That(items[1].ErrorColumnIndex, Is.EqualTo(1));
        }

        [Test]
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
            Assert.That(items[0].Value.DoubleProperty, Is.EqualTo(int1));
            Assert.That(items[0].Value.Int32Property, Is.EqualTo(Math.Round(dou1)));
            Assert.That(items[0].Value.StringProperty, Is.EqualTo(str1));
            Assert.That(items[1].Value.DoubleProperty, Is.EqualTo(default(double)));
            Assert.That(items[1].Value.Int32Property, Is.EqualTo(default(int)));
            Assert.That(items[1].Value.StringProperty, Is.EqualTo(str2));
        }

        [Test]
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
            Assert.That(items[0].Value.EnumProperty, Is.EqualTo(SampleEnum.Value1));
            Assert.That(items[1].Value.EnumProperty, Is.EqualTo(SampleEnum.Value2));
            Assert.That(items[2].Value.EnumProperty, Is.EqualTo(SampleEnum.Value3));
        }

        [Test]
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
            var obj = mapper.Take<SampleClass>().Select(o => o.Value).ToArray()[0];

            // Assert
            Assert.That(obj.StringProperty, Is.EqualTo("a"));
            Assert.That(obj.GeneralProperty, Is.EqualTo("b"));
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void Take_WithFirstRowIndex_ShouldImportExpectedRows(bool hasHeader)
        {
            // Arrange
            const int firstRowIndex = 100;
            const string sheetName = "sheet2";
            var workbook = GetSimpleWorkbook(DateTime.Now, "a");
            const string nameString = "StringProperty";
            const string nameGeneral = "GeneralProperty";
            var sheet = workbook.GetSheet(sheetName);

            if (hasHeader)
            {
                var headerRow = sheet.CreateRow(firstRowIndex);
                headerRow.CreateCell(0).SetCellValue(nameGeneral);
                headerRow.CreateCell(1).SetCellValue(nameString);
            }

            var firstDataRowIndex = hasHeader ? firstRowIndex + 1 : firstRowIndex;
            var row1 = sheet.CreateRow(firstDataRowIndex);
            row1.CreateCell(0).SetCellValue("a");
            row1.CreateCell(1).SetCellValue("b");
            var row2 = sheet.CreateRow(firstDataRowIndex + 1);
            row2.CreateCell(0).SetCellValue("c");
            row2.CreateCell(1).SetCellValue("d");

            var mapper = new Mapper(workbook) { HasHeader = hasHeader, FirstRowIndex = firstRowIndex };
            mapper.Map<SampleClass>(0, o => o.GeneralProperty);
            mapper.Map<SampleClass>(1, o => o.StringProperty);

            // Act
            var obj = mapper.Take<SampleClass>(sheetName).ToList();

            // Assert
            Assert.That(obj.Count, Is.EqualTo(2));
            Assert.That(obj[0].Value.GeneralProperty, Is.EqualTo("a"));
            Assert.That(obj[0].Value.StringProperty, Is.EqualTo("b"));
            Assert.That(obj[1].Value.GeneralProperty, Is.EqualTo("c"));
            Assert.That(obj[1].Value.StringProperty, Is.EqualTo("d"));
        }

        private class TestTrimClass
        {
            public string StringProperty { get; set; }
        }

        [Test]
        public void Take_MapByNameAndExtraSpaceInExcelColumnName_MapsAsTrimmed()
        {
            // Arrange
            const string str1 = "aBC";
            const string str2 = "BCD";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);

            // Header row with extra spaces
            sheet.GetRow(0).CreateCell(0).SetCellValue(" Name  ");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(str1);

            // Row #2
            sheet.GetRow(2).CreateCell(0).SetCellValue(str2);
            var mapper = new Mapper(workbook);
            mapper.Map<TestTrimClass>("Name", o => o.StringProperty);

            // Act
            var items = mapper.Take<TestTrimClass>().ToList();

            // Assert
            Assert.That(items[0].Value.StringProperty, Is.EqualTo(str1));
            Assert.That(items[1].Value.StringProperty, Is.EqualTo(str2));
        }

        private class TestGuidClass
        {
            public Guid ID { get; set; }
        }

        [Test]
        public void Take_GuidColumn_ParseAndSetGuid()
        {
            // Arrange
            var id = Guid.NewGuid();
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            // Header row with extra spaces
            sheet.GetRow(0).CreateCell(0).SetCellValue("ID");

            // Row #1
            sheet.GetRow(1).CreateCell(0).SetCellValue(id.ToString());

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<TestGuidClass>().ToList();

            // Assert
            Assert.That(items[0].Value.ID, Is.EqualTo(id));
        }

        [Test]
        public void Take_ColumnName_CaseInsensitive()
        {
            // Arrange
            const string value = "dummy";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            sheet.GetRow(0).CreateCell(0).SetCellValue(nameof(TestClass.String).ToUpperInvariant());
            sheet.GetRow(1).CreateCell(0).SetCellValue(value);

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<TestClass>().ToList();

            // Assert
            Assert.That(items[0].Value.String, Is.EqualTo(value));
        }

        [Test]
        public void Take_SkipHiddenRows_True()
        {
            // Arrange
            const string value = "dummy";
            const string hiddenValue = "hidden dummy";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);

            sheet.GetRow(0).CreateCell(0).SetCellValue(nameof(TestClass.String));
            sheet.GetRow(1).CreateCell(0).SetCellValue(value);
            sheet.GetRow(2).CreateCell(0).SetCellValue(hiddenValue);
            sheet.GetRow(2).Hidden = true;

            var mapper = new Mapper(workbook)
            {
                SkipHiddenRows = true,
            };

            // Act
            var items = mapper.Take<TestClass>().ToList();

            // Assert
            Assert.That(items.Count, Is.EqualTo(1));
            Assert.That(items[0].Value.String, Is.EqualTo(value));
        }

        [Test]
        public void Take_SkipHiddenRows_False()
        {
            // Arrange
            const string value = "dummy";
            const string hiddenValue = "hidden dummy";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);

            sheet.GetRow(0).CreateCell(0).SetCellValue(nameof(TestClass.String));
            sheet.GetRow(1).CreateCell(0).SetCellValue(value);
            sheet.GetRow(2).CreateCell(0).SetCellValue(hiddenValue);
            sheet.GetRow(2).Hidden = true;

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<TestClass>().ToList();

            // Assert
            Assert.That(items.Count, Is.EqualTo(2));
            Assert.That(items[0].Value.String, Is.EqualTo(value));
            Assert.That(items[1].Value.String, Is.EqualTo(hiddenValue));
        }

        [Test]
        public void Take_DateTime_And_DateTimeOffice()
        {
            // Arrange
            var value = DateTimeOffset.Now.Truncate();
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);

            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(TestClass.DateTime));
            header.CreateCell(1).SetCellValue(nameof(TestClass.DateTimeOffsetProperty));

            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue(value.ToString(CultureInfo.InvariantCulture));
            row1.CreateCell(1).SetCellValue(value.ToString(CultureInfo.InvariantCulture));

            var row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellValue(value.DateTime);
            row2.CreateCell(1).SetCellValue(value.DateTime);

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<TestClass>().ToList();

            // Assert
            Assert.That(items[0].Value.DateTime, Is.EqualTo(value.DateTime));
            Assert.That(items[0].Value.DateTimeOffsetProperty, Is.EqualTo(value));
            Assert.That(items[1].Value.DateTime, Is.EqualTo(value.DateTime));
            Assert.That(items[1].Value.DateTimeOffsetProperty, Is.EqualTo(value));
        }

        [Test]
        public void Take_Nullable_DateTime_And_DateTimeOffice()
        {
            // Arrange
            var value = DateTimeOffset.Now.Truncate();
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);

            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(NullableClass.NormalString));
            header.CreateCell(1).SetCellValue(nameof(NullableClass.NullableDateTime));
            header.CreateCell(2).SetCellValue(nameof(NullableClass.NullableDateTimeOffset));

            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue(value.ToString(CultureInfo.InvariantCulture));
            row1.CreateCell(1).SetCellValue(value.ToString(CultureInfo.InvariantCulture));
            row1.CreateCell(2).SetCellValue(value.ToString(CultureInfo.InvariantCulture));

            var row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellValue(value.ToString(CultureInfo.InvariantCulture));
            row2.CreateCell(1).SetCellValue(value.DateTime);
            row2.CreateCell(2).SetCellValue(value.DateTime);

            var row3 = sheet.CreateRow(3);
            row3.CreateCell(0).SetCellValue(value.ToString(CultureInfo.InvariantCulture));
            row3.CreateCell(1).SetCellValue(default(string));
            row3.CreateCell(2).SetCellValue(default(string));

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<NullableClass>().ToList();

            // Assert
            Assert.That(items[0].Value.NullableDateTime, Is.EqualTo(value.DateTime));
            Assert.That(items[0].Value.NullableDateTimeOffset, Is.EqualTo(value));
            Assert.That(items[1].Value.NullableDateTime, Is.EqualTo(value.DateTime));
            Assert.That(items[1].Value.NullableDateTimeOffset, Is.EqualTo(value));
            Assert.That(items[2].Value.NormalString, Is.EqualTo(value.ToString(CultureInfo.InvariantCulture)));
            Assert.That(items[2].Value.NullableDateTime, Is.Null);
            Assert.That(items[2].Value.NullableDateTimeOffset, Is.Null);
        }

        [Test]
        public void Take_NumericCell_As_String()
        {
            // Arrange
            var doubleValue = 35.3456d;
            var dateValue = DateTimeOffset.Now.Truncate();
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var dateCellStyle = workbook.CreateCellStyle();
            dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyyMMdd");
            var doubleCellStyle = workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("$0.00\" Surplus\";$-0.00\" Shortage\"");
            var dataFormatter = new DataFormatter();
            var fe = new XSSFFormulaEvaluator(workbook);

            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(TestClass.String));

            var row1 = sheet.CreateRow(1);
            var cell = row1.CreateCell(0);
            cell.SetCellValue(doubleValue.ToString(CultureInfo.InvariantCulture));

            var row2 = sheet.CreateRow(2);
            cell = row2.CreateCell(0);
            cell.CellStyle = dateCellStyle;
            cell.SetCellValue(dateValue.DateTime);
            var dateString = dataFormatter.FormatCellValue(cell, fe);

            var row3 = sheet.CreateRow(3);
            cell = row3.CreateCell(0);
            cell.CellStyle = doubleCellStyle;
            cell.SetCellValue(doubleValue);
            var doubleString = dataFormatter.FormatCellValue(cell, fe); // $35.35 Surplus

            var mapper = new Mapper(workbook);

            // Act
            var items = mapper.Take<TestClass>().ToList();

            // Assert
            Assert.That(items[0].Value.String, Is.EqualTo(doubleValue.ToString(CultureInfo.InvariantCulture)));
            Assert.That(items[1].Value.String, Is.EqualTo(dateString));
            Assert.That(items[2].Value.String, Is.EqualTo(doubleString));
        }
    }
}
