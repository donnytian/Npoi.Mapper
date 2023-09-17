using System;
using System.IO;
using System.Linq;
using Npoi.Mapper;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace test
{
    [TestFixture]
    public class ImportDynamicTests : TestBase
    {
        [Test]
        public void TakeDynamic_Possitive()
        {
            // Arrange
            var boolProperty = "  "; // Given a invalid property name, mapper should populate property with name according the column index. e.g. A, B, AC.
            var dateProperty = "ColumnDate";
            var stringProperty = "Column String";
            var date1 = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(703).SetCellValue(boolProperty); // Column AAB in Excel.
            header.CreateCell(5).SetCellValue(dateProperty);
            header.CreateCell(10).SetCellValue(stringProperty);
            var row = sheet.CreateRow(1);
            row.CreateCell(703).SetCellValue(true);
            var dateCell = row.CreateCell(5);
            dateCell.SetCellValue(date1);
            // Format cell as date time to ensure the mapper can infer it as DateTime type since date time is store as double in Excel.
            dateCell.CellStyle = MapHelper.CreateCellStyle(workbook, "dd-MM-yyyy hh:mm:ss");
            row.CreateCell(10).SetCellValue(str1);

            // Act
            var mapper = new Mapper(workbook);
            //mapper.Save(new FileStream("dddd.xlsx", FileMode.Create)); // Use this to lookup the column name (like AAB) in Excel...
            var objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.AreEqual(date1.ToLongDateString(), objs[0].Value.ColumnDate.ToLongDateString());
            Assert.AreEqual(str1, objs[0].Value.ColumnString);
            Assert.IsTrue(objs[0].Value.AAB);
        }

        [Test]
        public void TakeDynamic_LookupColumnType()
        {
            // Arrange
            var boolProperty = "  "; // Given a invalid property name, mapper should populate property with name according the column index. e.g. A, B, AC.
            var dateProperty = "ColumnDate";
            var stringProperty = "Column String";
            var date1 = DateTime.Now;
            const string str1 = "aBC";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(703).SetCellValue(boolProperty); // Column AAB in Excel.
            header.CreateCell(5).SetCellValue(dateProperty);
            header.CreateCell(10).SetCellValue(stringProperty);
            var row = sheet.CreateRow(1); // objs[0]
            row = sheet.CreateRow(5);     // objs[1]
            row = sheet.CreateRow(6);     // objs[2]
            row = sheet.CreateRow(10);    // objs[3]
            row.CreateCell(703).SetCellValue(true);
            var dateCell = row.CreateCell(5);
            dateCell.SetCellValue(date1);
            // Format cell as date time to ensure the mapper can infer it as DateTime type since date time is store as double in Excel.
            dateCell.CellStyle = MapHelper.CreateCellStyle(workbook, "dd-MM-yyyy hh:mm:ss");
            row.CreateCell(10).SetCellValue(str1);

            // Act
            var mapper = new Mapper(workbook);
            var objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.AreEqual(date1.ToLongDateString(), objs[3].Value.ColumnDate.ToLongDateString());
            Assert.AreEqual(str1, objs[3].Value.ColumnString);
            Assert.IsTrue(objs[3].Value.AAB);
        }

        [Test]
        public void TakeDynamic_Modify_ThenExport()
        {
            // Arrange
            var tempFileName = "TakeDynamic_Modify_ThenExport.xlsx";
            var dateProperty = "ColumnDate";
            var date1 = DateTime.Now;
            var date2 = date1.AddMonths(1);
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(5).SetCellValue(dateProperty);
            var row = sheet.CreateRow(5);
            var dateCell = row.CreateCell(5);
            dateCell.SetCellValue(date1);
            // Format cell as date time to ensure the mapper can infer it as DateTime type since date time is store as double in Excel.
            dateCell.CellStyle = MapHelper.CreateCellStyle(workbook, "dd-MM-yyyy hh:mm:ss");

            // Act
            var mapper = new Mapper(workbook);
            var objs = mapper.Take<dynamic>().ToList();
            objs[0].Value.ColumnDate = date2;
            if (File.Exists(tempFileName)) File.Delete(tempFileName);
            mapper.Put(new[] { objs[0].Value });
            mapper.Save(new FileStream(tempFileName, FileMode.Create), false);

            mapper = new Mapper(tempFileName);
            objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.AreEqual(date2.ToLongDateString(), objs[0].Value.ColumnDate.ToLongDateString());
            Assert.AreEqual(164, mapper.Workbook.GetSheetAt(0).GetRow(1).GetCell(5).CellStyle.DataFormat);
            File.Delete(tempFileName);
        }

        [Test]
        public void TakeDynamic_IgnoredChars_Issue7()
        {
            // Arrange
            var str = "dummy";
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(5).SetCellValue("N.I?F@");
            var row = sheet.CreateRow(5);
            var dateCell = row.CreateCell(5);
            dateCell.SetCellValue(str);

            // Act
            var mapper = new Mapper(workbook);
            var objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.AreEqual(str, objs[0].Value.NIF);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void TakeDynamic_WithFirstRowIndex_ShouldImportExpectedRows(bool hasHeader)
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

            // Act
            var obj = mapper.Take<dynamic>(sheetName).ToList();

            // Assert
            Assert.AreEqual(2, obj.Count);
            if (hasHeader)
            {
                Assert.AreEqual("a", obj[0].Value.GeneralProperty);
                Assert.AreEqual("b", obj[0].Value.StringProperty);
                Assert.AreEqual("c", obj[1].Value.GeneralProperty);
                Assert.AreEqual("d", obj[1].Value.StringProperty);
            }
            else
            {
                Assert.AreEqual("a", obj[0].Value.A);
                Assert.AreEqual("b", obj[0].Value.B);
                Assert.AreEqual("c", obj[1].Value.A);
                Assert.AreEqual("d", obj[1].Value.B);
            }
        }

        [Test]
        public void TakeDynamicWithColumnType_With_TypeResolver()
        {
            // Arrange
            const string stringValue = "dummy";
            const int intValue = 11;
            const double doubleValue = 1.11d;
            var dateTimeValue = DateTime.Now.Truncate();
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("A"); // int
            header.CreateCell(1).SetCellValue("B"); // DateTime
            header.CreateCell(2).SetCellValue("C"); // string
            header.CreateCell(3).SetCellValue("D"); // double - auto-detected
            var row1 = sheet.CreateRow(1); // put invalid types in row1
            row1.CreateCell(0).SetCellValue("invalid int value");
            row1.CreateCell(1).SetCellValue("invalid date time value");
            row1.CreateCell(2).SetCellValue(dateTimeValue);
            row1.CreateCell(3).SetCellValue(intValue);
            var row2 = sheet.CreateRow(2); // put valid values in row2
            row2.CreateCell(0).SetCellValue(intValue);
            row2.CreateCell(1).SetCellValue(dateTimeValue);
            row2.CreateCell(2).SetCellValue(stringValue);
            row2.CreateCell(3).SetCellValue(doubleValue);

            Type TypeResolver(ICell cell) => cell.ColumnIndex switch
            {
                0 => typeof(int),
                1 => typeof(DateTime),
                2 => typeof(string),
                _ => null, // return null to let mapper detect from the first data row.
            };

            // Act
            var mapper = new Mapper(workbook);
            var objs = mapper.TakeDynamicWithColumnType(TypeResolver).ToList();

            // Assert
            Assert.AreEqual(0, objs[0].ErrorColumnIndex);
            Assert.AreEqual(intValue, objs[0].Value.D);
            Assert.AreEqual(intValue, objs[1].Value.A);
            Assert.AreEqual(dateTimeValue, objs[1].Value.B);
            Assert.AreEqual(stringValue, objs[1].Value.C);
            Assert.AreEqual(doubleValue, objs[1].Value.D);
        }

                [Test]
        public void TakeDynamic_TwoSheets_WithSameHeaderName()
        {
            // Arrange
            const string stringValue = "dummy";
            const int intValue = 11;
            const double doubleValue = 4.21d;
            var dateTimeValue = DateTime.Now.Truncate();

            var workbook = GetEmptyWorkbook();

            var sheet1 = workbook.CreateSheet("sheet1");
            var header = sheet1.CreateRow(0);
            header.CreateCell(0).SetCellValue("A");
            header.CreateCell(1).SetCellValue("sheet1B");

            var row1 = sheet1.CreateRow(1);
            row1.CreateCell(0).SetCellValue(stringValue);
            row1.CreateCell(1).SetCellValue(intValue);

            var sheet2 = workbook.CreateSheet("sheet2");
            header = sheet2.CreateRow(0);
            header.CreateCell(0).SetCellValue("A");
            header.CreateCell(1).SetCellValue("sheet2B");

            row1 = sheet2.CreateRow(1);
            row1.CreateCell(0).SetCellValue(doubleValue);
            row1.CreateCell(1).SetCellValue(dateTimeValue);

            // Act
            var mapper = new Mapper(workbook);
            var objs1 = mapper.Take<dynamic>("sheet1").ToList();
            var objs2 = mapper.Take<dynamic>("sheet2").ToList();

            // Assert
            Assert.AreEqual(stringValue, objs1[0].Value.A);
            Assert.AreEqual(intValue, objs1[0].Value.sheet1B);
            Assert.AreEqual(doubleValue, objs2[0].Value.A);
            var diff = dateTimeValue.ToOADate() - objs2[0].Value.sheet2B;
            const double epsilon = 0.0000000001;
            Assert.IsTrue(Math.Abs(diff) <  epsilon);
        }
    }
}
