using System;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies that a property is the container (ICollection) to receive values cross multiple columns.
    /// For example a Values property can receive items in column "Value1", "Value2" and etc.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class MultiColumnContainerAttribute : FieldAttribute
    {
        /// <summary>
        /// Initialize a new instance of <see cref="MultiColumnContainerAttribute"/> class.
        /// </summary>
        public MultiColumnContainerAttribute(Type columnResolverType) : base(columnResolverType)
        {
        }
    }
}
