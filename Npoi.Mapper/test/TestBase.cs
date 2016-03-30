using System.IO;
using NPOI.SS.UserModel;

namespace test
{
    /// <summary>
    /// Base class for test classes.
    /// </summary>
    public abstract class TestBase
    {
        protected Stream InputWorkbookStream { get; set; }

        protected IWorkbook Workbook { get; set; }
    }
}
