using System;
using System.Collections.Generic;
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
    public class ImporterTests : TestBase
    {
        [TestInitialize]
        public void InitializeTest()
        {
        }

        [TestCleanup]
        public void CleanupTest()
        {
            InputWorkbookStream?.Dispose();
            Workbook = null;
        }

        [TestMethod]
        public void ImporterConstructorTest()
        {
            // Prepare
            InputWorkbookStream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Importer(InputWorkbookStream);

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ImporterConstructorNullExceptionTest()
        {
            // Prepare
            Stream nullStream = null;

            // Act
            // ReSharper disable once ExpressionIsAlwaysNull
            // ReSharper disable once UnusedVariable
            var importer = new Importer(nullStream);


            // Assert
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ImporterConstructorNullExceptionTest2()
        {
            // Prepare
            IWorkbook nullWorkbook = null;

            // Act
            // ReSharper disable once ExpressionIsAlwaysNull
            // ReSharper disable once UnusedVariable
            var importer = new Importer(nullWorkbook);


            // Assert
        }

        [TestMethod]
        public void TakeByHeaderIndexTest()
        {
            // Prepare
            var date = DateTime.Now;
            var str = "aBC";
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("sheet1");
            var sheet = workbook.CreateSheet("sheet2");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("DateProperty");
            header.CreateCell(1).SetCellValue("StringProperty");
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(date);
            row.CreateCell(1).SetCellValue(str);
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>(1).ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(1, objs.Count);

            var obj = objs[0];
            var objDate = obj.Value.DateProperty;

            Assert.AreEqual(date.ToLongTimeString(), objDate.ToLongTimeString());
            Assert.AreEqual(str, obj.Value.StringProperty);
        }
    }
}
