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
        public void ImporterConstructorStreamTest()
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
        public void ImporterConstructorWorkbookTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.MaxValue, "dummy");

            // Act
            var importer = new Importer(workbook);

            // Assert
            Assert.IsNotNull(importer);
            Assert.IsNotNull(importer.Workbook);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ImporterConstructorNullStreamTest()
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
        public void ImporterConstructorNullWorkbookTest()
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
        public void ImporterNoElementTest()
        {
            // Prepare
            var workbook = new XSSFWorkbook();
            var header = workbook.CreateSheet("sheet1").CreateRow(0);
            header.CreateCell(0).SetCellValue("StringProperty");
            header.CreateCell(1).SetCellValue("Int32Property");
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>(0);

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count());
        }

        [TestMethod]
        public void ImporterEmptySheetTest()
        {
            // Prepare
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("sheet1");
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>(0);

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count());
        }

        [TestMethod]
        public void TakeByHeaderIndexTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>(1).ToList();

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
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Importer(workbook);

            // Act
            // ReSharper disable once UnusedVariable
            var objs = importer.TakeByHeader<SampleClass>(10).ToList();

            // Assert
        }

        [TestMethod]
        public void TakeByHeaderNameTest()
        {
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>("sheet2").ToList();

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
            // Prepare
            var date = DateTime.Now;
            const string str = "aBC";
            var workbook = GetSimpleWorkbook(date, str);
            var importer = new Importer(workbook);

            // Act
            var objs = importer.TakeByHeader<SampleClass>("notExistSheet").ToList();

            // Assert
            Assert.IsNotNull(objs);
            Assert.AreEqual(0, objs.Count);
        }
    }
}
