using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies to ignore a property for mapping.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class IgnoreAttribute : Attribute
    {
    }
}
