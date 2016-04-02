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
    public class ExportTests : TestBase
    {
        [TestMethod]
        public void SaveSheetTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save<SampleClass>("test.xlsx", 1);


            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }

        [TestMethod]
        public void SaveObjectsTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();

            // Act
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }

        [TestMethod]
        public void FormatAttributeTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.BuiltinFormatProperty = DateTime.Now;
            objs[0].Value.CustomFormatProperty = 100.234;

            // Act
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }

        [TestMethod]
        public void FormatMethodTest()
        {
            // Prepare
            var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
            var exporter = new Mapper(workbook);
            var objs = exporter.Take<SampleClass>(1).ToList();
            objs[0].Value.DateProperty = DateTime.Now;
            objs[0].Value.DoubleProperty = 100.234;

            // Act
            exporter.Format<SampleClass>(0xf, o => o.DateProperty);
            exporter.Format<SampleClass>("0%", o => o.DoubleProperty);
            exporter.Save("test.xlsx", objs.Select(info => info.Value), "newSheet");

            // Assert
            Assert.IsNotNull(objs);
            Assert.IsNotNull(exporter);
            Assert.IsNotNull(exporter.Workbook);
        }
    }
}
