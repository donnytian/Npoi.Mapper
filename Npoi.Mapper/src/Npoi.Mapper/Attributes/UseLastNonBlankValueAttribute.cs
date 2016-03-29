using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property should use the last non-blank value.
    /// Typically handle the value blank issue in merged cells.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class UseLastNonBlankValueAttribute : Attribute
    {
        /// <summary>
        /// Initialize a new instance of <see cref="UseLastNonBlankValueAttribute"/> class.
        /// </summary>
        public UseLastNonBlankValueAttribute()
        {
        }
    }
}
