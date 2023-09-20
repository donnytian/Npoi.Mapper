using System;
using System.Diagnostics;
using System.Linq;
using Npoi.Mapper;
using NUnit.Framework;

namespace test;

[TestFixture]
public class PerformanceTests : TestBase
{
    [Test]
    [TestCase(1_000_000)]     // 2224 ms vs 2634 ms, expression tree improved 10%-20% comparing to the reflection way.
    [Ignore("Do not run this long-running test unless you really understand what you are going to do.")]
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
        TestContext.Out.WriteLine($"Total Row:{count:0000000} - {watch.Elapsed}");
        Assert.AreEqual(count, objs.Count);
    }
}
