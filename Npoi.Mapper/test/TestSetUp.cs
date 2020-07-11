using System;
using NUnit.Framework;

namespace test
{
    /// <summary>
    /// Contains setup and teardown for all text fixtures in a given namespace.
    /// </summary>
    [SetUpFixture]
    public class TestSetUp
    {
        [OneTimeSetUp]
        public void RunBeforeAnyTests()
        {
            // NUnit3 changed the working directory to the folder of runner instead of the output (bing/Debug).
            // So here we set it back.
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
        }
    }
}
