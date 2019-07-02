using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Npoi.Mapper;

namespace test
{
    [TestClass]
    public class PerformanceTests : TestBase
    {
        [TestMethod]
        [DataRow(100)]           // 37 ms
        [DataRow(10_000)]       // 71 ms
        [DataRow(1_000_000)]    // 7284 ms
        public void TakeDynamic_Performance_Tests(int count)
        {
            // Arrange
            var watch = new Stopwatch();
            var now = DateTime.Now;
            var workbook = GetBlankWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("string");
            header.CreateCell(1).SetCellValue("int");
            header.CreateCell(2).SetCellValue("date");

            for (var i = 1; i <= count; i++)
            {
                var row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue("this is a dummy string!");
                row.CreateCell(1).SetCellValue(i);
                row.CreateCell(2).SetCellValue(now.AddSeconds(i));
            }

            var mapper = new Mapper(workbook);

            // Act
            watch.Start();
            var objs = mapper.Take<dynamic>().ToList();
            watch.Stop();

            // Assert
            Trace.WriteLine($"Total Row:{count:0000000} - {watch.ElapsedMilliseconds} ms");
            Assert.AreEqual(count, objs.Count);
        }
    }
}
