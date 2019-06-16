﻿using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace test
{
    /// <summary>
    /// Base class for test classes.
    /// </summary>
    public abstract class TestBase
    {
        protected Stream InputWorkbookStream { get; set; }

        protected IWorkbook Workbook { get; set; }

        #region Protected Methods

        protected static IWorkbook GetSimpleWorkbook(DateTime dateValue, string stringValue)
        {
            var workbook = GetEmptyWorkbook();
            workbook.CreateSheet("sheet1");
            var sheet = workbook.CreateSheet("sheet2");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("DateProperty");
            header.CreateCell(1).SetCellValue("StringProperty");
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(dateValue);
            row.CreateCell(1).SetCellValue(stringValue);

            return workbook;
        }

        protected static IWorkbook GetBlankWorkbook()
        {
            var workbook = GetEmptyWorkbook();
            workbook.CreateSheet("sheet1");

            return workbook;
        }

        protected static IWorkbook GetEmptyWorkbook()
        {
            var workbook = new XSSFWorkbook();

            return workbook;
        }

        protected static IWorkbook WriteAndReadBack(IWorkbook workbook, string fileName = "TempWrite")
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            return WorkbookFactory.Create(fileName);
        }

        protected static void CreateShiftedRowsWorkbook(string fromFile, string toFile, string sheetName, int shiftRows)
        {
            IWorkbook book;
            using (var stream = new FileStream(fromFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                book = new XSSFWorkbook(stream);
            }

            book.GetSheet(sheetName).ShiftRows(0, shiftRows, shiftRows);
            book.GetSheet(sheetName).CreateRow(0).CreateCell(0).SetCellValue("Ignored");
            using (var fileStream = new FileStream(toFile, FileMode.Create))
            {
                book.Write(fileStream);
            }

            book.Close();
        }
        #endregion
    }
}
