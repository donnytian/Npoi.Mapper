using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper.Attributes;

namespace test.Sample
{
    /// <summary>
    /// The base class for sample classes.
    /// </summary>
    public class BaseClass
    {
        public string BaseStringProperty { get; set; }

        [Ignore]
        public string BaseIgnoredProperty { get; set; }
    }
}
