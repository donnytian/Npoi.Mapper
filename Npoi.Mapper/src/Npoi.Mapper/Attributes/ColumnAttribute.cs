using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property is mapped to a column by index.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : FieldAttribute
    {
        /// <summary>
        /// Column index.
        /// </summary>
        public int Index { get; } = -1;

        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnAttribute"/> class.
        /// </summary>
        public ColumnAttribute(int index, Type columnResolverType = null) : base(columnResolverType)
        {
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index));

            Index = index;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnAttribute"/> class.
        /// </summary>
        public ColumnAttribute(string name, Type columnResolverType = null) : base(columnResolverType)
        {
            Name = name;
        }
    }
}
