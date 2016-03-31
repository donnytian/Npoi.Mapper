using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property is ignored.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnoreAttribute : Attribute
    {
    }
}
