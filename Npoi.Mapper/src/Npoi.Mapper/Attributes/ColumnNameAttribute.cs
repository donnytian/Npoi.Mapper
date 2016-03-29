using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property is mapped to a column by name.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnNameAttribute : FieldAttribute
    {
        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnNameAttribute"/> class.
        /// </summary>
        public ColumnNameAttribute(string name, Type columnResolverType = null) : base(columnResolverType)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Name = name;
        }
    }
}
