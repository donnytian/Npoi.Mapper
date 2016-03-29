using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Base class for attributes that can apply to object field.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field)]
    public abstract class FieldAttribute : Attribute
    {
        /// <summary>
        /// The type of class that is derived from <see cref="ColumnResolver{TTarget}"/> class.
        /// </summary>
        public Type ColumnResolverType { get; }

        /// <summary>
        /// Initialize a new instance of <see cref="FieldAttribute"/> class.
        /// </summary>
        /// <param name="columnResolverType"></param>
        protected FieldAttribute(Type columnResolverType)
        {
            ColumnResolverType = columnResolverType;
        }
    }
}
