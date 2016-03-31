using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required when map from property to column.
    /// </summary>
    public class MappingInfo
    {
        #region Properties

        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; protected set; }

        /// <summary>
        /// Column index.
        /// </summary>
        public int Index { get; protected set; } = -1;

        /// <summary>
        /// Mapped property for this column value according column name.
        /// </summary>
        public PropertyInfo Property { get; protected set; }

        /// <summary>
        /// The column resolver type.
        /// </summary>
        public Type ResolverType { get; set; }

        /// <summary>
        /// Indicate whether to use the last non-blank value.
        /// Typically handle the blank error in merged cells.
        /// </summary>
        public bool UseLastNonBlankValue { get; set; }

        /// <summary>
        /// Indicate whether to ignore the property.
        /// </summary>
        public bool IgnoreProperty { get; set; }

        /// <summary>
        /// Indicate whether the mapping already done or not.
        /// </summary>
        public bool Mapped { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="MappingInfo"/> class.
        /// </summary>
        /// <param name="name">The column name</param>
        /// <param name="property">The target property info</param>
        /// <param name="resolverType">The column resolver type</param>
        public MappingInfo(string name, PropertyInfo property, Type resolverType = null)
        {
            Name = name;
            Property = property;
            ResolverType = resolverType;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="MappingInfo"/> class.
        /// </summary>
        /// <param name="index">The column index</param>
        /// <param name="property">The target property info</param>
        /// <param name="resolverType">The column resolver type</param>
        public MappingInfo(int index, PropertyInfo property, Type resolverType = null)
        {
            Index = index;
            Property = property;
            ResolverType = resolverType;
        }

        #endregion
    }
}
