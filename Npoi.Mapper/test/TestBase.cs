using System;
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

        /// <summary>
        /// Gets a workbook with 2 sheets("sheet1" and "sheet2") and with 2 rows in "sheet2":
        /// <para>DateProperty StringProperty</para>
        /// <para>dateValue    stringValue</para>
        /// </summary>
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

        /// <summary>
        /// Gets a workbook with a empty sheet named "sheet1".
        /// </summary>
        /// <returns></returns>
        protected static IWorkbook GetBlankWorkbook()
        {
            var workbook = GetEmptyWorkbook();
            workbook.CreateSheet("sheet1");

            return workbook;
        }

        /// <summary>
        /// Gets a workbook without any sheets and rows.
        /// </summary>
        /// <returns></returns>
        protected static IWorkbook GetEmptyWorkbook()
        {
            var workbook = new XSSFWorkbook();

            return workbook;
        }

        protected static IWorkbook WriteAndReadBack(IWorkbook workbook, string fileName = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = Guid.NewGuid().ToString();
            }

            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs, true);
            }

            IWorkbook newWorkbook;

            try
            {
                newWorkbook = WorkbookFactory.Create(fileName);
            }
            finally
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }

            return newWorkbook;
        }
    }
}
