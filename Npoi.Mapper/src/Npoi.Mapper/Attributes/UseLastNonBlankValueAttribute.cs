using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies to use the last non-blank value when reading from cells for this property.
    /// Typically handle the blank error in merged cells.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class UseLastNonBlankValueAttribute : Attribute
    {
    }
}
