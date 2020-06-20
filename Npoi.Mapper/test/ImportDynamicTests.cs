using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;

namespace test
{
    [TestClass]
    public class ImportDynamicTests : TestBase
    {
        [TestMethod]
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

        [TestMethod]
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

        [TestMethod]
        public void TakeDynamic_Modify_ThenExport()
        {
            // Arrange
            var tempFileName = "_tempFile.xlsx";
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
            mapper.Save(new FileStream(tempFileName, FileMode.Create));

            mapper = new Mapper(tempFileName);
            objs = mapper.Take<dynamic>().ToList();

            // Assert
            Assert.AreEqual(date2.ToLongDateString(), objs[0].Value.ColumnDate.ToLongDateString());
            Assert.AreEqual(164, mapper.Workbook.GetSheetAt(0).GetRow(1).GetCell(5).CellStyle.DataFormat);
        }

        [TestMethod]
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

        [TestMethod]
        [DataRow(true)]
        [DataRow(false)]
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
    }
}
