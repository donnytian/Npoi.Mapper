using System;
using System.Reflection;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required for a target property.
    /// </summary>
    public class PropertyMeta
    {
        #region Properties

        /// <summary>
        /// Column name.
        /// </summary>
        public string ColumnName { get; protected set; }

        /// <summary>
        /// Column index.
        /// </summary>
        public int ColumnIndex { get; protected set; } = -1;

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
        public bool Ignored { get; set; }

        /// <summary>
        /// Indicate whether the mapping was setup or not.
        /// </summary>
        public bool Mapped { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="PropertyMeta"/> class.
        /// </summary>
        /// <param name="property">The target property info</param>
        /// <param name="resolverType">The column resolver type</param>
        public PropertyMeta(PropertyInfo property, Type resolverType = null)
        {
            Property = property;
            ResolverType = resolverType;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="PropertyMeta"/> class.
        /// </summary>
        /// <param name="columnName">The column name</param>
        /// <param name="property">The target property info</param>
        /// <param name="resolverType">The column resolver type</param>
        public PropertyMeta(string columnName, PropertyInfo property, Type resolverType = null)
        {
            ColumnName = columnName;
            Property = property;
            ResolverType = resolverType;
        }

        /// <summary>
        /// Initialize a new instance of <see cref="PropertyMeta"/> class.
        /// </summary>
        /// <param name="columnIndex">The column index</param>
        /// <param name="property">The target property info</param>
        /// <param name="resolverType">The column resolver type</param>
        public PropertyMeta(int columnIndex, PropertyInfo property, Type resolverType = null)
        {
            ColumnIndex = columnIndex;
            Property = property;
            ResolverType = resolverType;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// 
        /// </summary>
        /// <param name="target"></param>
        /// <param name="resolverType"></param>
        public static void MapColumn(PropertyMeta target, Type resolverType)
        {
            if (target != null)
            {
                target.ResolverType = resolverType;
                target.Ignored = false;
                //mapping.Mapped = true;
            }
        }

        #endregion

        #region Private Methods



        #endregion
    }
}
