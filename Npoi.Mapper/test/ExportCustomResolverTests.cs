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
    public class ExportCustomResolverTests : TestBase
    {
        [TestMethod]
        public void ForHeader_ChangeHeaderStyle_ShouldChanged()
        {
            // Arrange
            const string str1 = "aBc";
            var workbook = GetBlankWorkbook();
            var row1 = workbook.GetSheetAt(0).CreateRow(0);
            row1.CreateCell(11).SetCellValue("StringProperty");

            var mapper = new Mapper(workbook);

            // Act
            mapper.ForHeader(cell =>
            {
                if (cell.ColumnIndex == 11)
                {
                    var style = cell.Sheet.Workbook.CreateCellStyle();
                    style.LeftBorderColor = 120;
                    cell.CellStyle = style;
                }
            });
            mapper.Put(new[] { new SampleClass { StringProperty = str1 } });

            // Assert
            var row2 = workbook.GetSheetAt(0).GetRow(1);
            Assert.AreEqual(120, row1.GetCell(11).CellStyle.LeftBorderColor);
            Assert.AreEqual(str1, row2.GetCell(11).StringCellValue);

        }

        [TestMethod]
        public void ForHeader_ChangeHeaderStyleForGivenHeaderIndex_ShouldChanged()
        {
            // Arrange
            const string str1 = "aBc";
            const int headerRowIndex = 5;
            var workbook = GetBlankWorkbook();

            var row1 = workbook.GetSheetAt(0).CreateRow(headerRowIndex);
            row1.CreateCell(11).SetCellValue("StringProperty");

            var mapper = new Mapper(workbook);
            mapper.HeaderRowIndex = headerRowIndex;
            // Act
            mapper.ForHeader(cell =>
            {
                if (cell.ColumnIndex == 11)
                {
                    var style = cell.Sheet.Workbook.CreateCellStyle();
                    style.LeftBorderColor = 120;
                    cell.CellStyle = style;
                }
            });
            mapper.Put(new[] { new SampleClass { StringProperty = str1 } });

            // Assert
            var row2 = workbook.GetSheetAt(0).GetRow(1 + mapper.HeaderRowIndex);
            Assert.AreEqual(120, row1.GetCell(11).CellStyle.LeftBorderColor);
            Assert.AreEqual(str1, row2.GetCell(11).StringCellValue);

        }
    }
}
