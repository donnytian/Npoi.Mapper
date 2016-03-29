using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property is mapped to a column by index.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnIndexAttribute : FieldAttribute
    {
        /// <summary>
        /// Column index.
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnNameAttribute"/> class.
        /// </summary>
        public ColumnIndexAttribute(int index, Type columnResolverType = null) : base(columnResolverType)
        {
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index));

            Index = index;
        }
    }
}
