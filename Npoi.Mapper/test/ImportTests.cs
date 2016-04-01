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
    public class ImportTests : TestBase
    {
        [TestInitialize]
        public void InitializeTest()
        {
        }

        [TestCleanup]
        public void CleanupTest()
        {
        }

        [TestMethod]
        public void ImporterConstructorStreamTest()
        {
            // Prepare
            var stream = new FileStream("Book1.xlsx", FileMode.Open);

            // Act
            var importer = new Mapper(stream);

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
            var importer = new Mapper(workbook);

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
            var importer = new Mapper(nullStream);


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
            var importer = new Mapper(nullWorkbook);

            // Assert
        }

        [TestMethod]
        public void ImporterConstructorFilePathTest()
        {
            // Prepare

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
            // Prepare

            // Act
            var importer = new Mapper("dummy.txt");

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
            // Prepare
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
            // Prepare
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
            // Prepare
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
            // Prepare
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
            // Prepare
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
    }
}
