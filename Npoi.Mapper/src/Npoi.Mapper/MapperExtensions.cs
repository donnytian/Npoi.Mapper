using System;
using System.Linq.Expressions;

namespace Npoi.Mapper
{
    /// <summary>
    /// Extension methods for <see cref="Mapper"/>.
    /// </summary>
    public static class MapperExtensions
    {
        /// <summary>
        /// Uses a custom format for all properties that have the same type.
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="propertyType">The type of property to format.</param>
        /// <param name="customFormat">The custom format for the specified type.</param>
        /// <returns>The <see cref="Mapper"/> itself.</returns>
        public static Mapper UseFormat(this Mapper mapper, Type propertyType, string customFormat)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (propertyType == null) throw new ArgumentNullException(nameof(propertyType));
            if (string.IsNullOrWhiteSpace(customFormat)) throw new ArgumentException($"Parameter '{nameof(customFormat)}' cannot be null or white space.");

            mapper.TypeFormats[propertyType] = customFormat;

            return mapper;
        }

        /// <summary>
        /// Map property to a column by specified column name and property name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertyName">The property name.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, string columnName, string propertyName, Type resolverType = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));

            var pi = typeof(T).GetProperty(propertyName, MapHelper.BindingFlag);

            if (pi == null) throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");

            return mapper.Map(columnName, pi, resolverType);
        }

        /// <summary>
        /// Map property to a column by specified column name and property selector.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, string columnName, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));

            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);

            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            return mapper.Map(columnName, pi, resolverType);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertyName">The property name.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, string propertyName, Type resolverType = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));

            var pi = typeof(T).GetProperty(propertyName, MapHelper.BindingFlag);

            if (pi == null) throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");

            return mapper.Map(columnIndex, pi, resolverType);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property selector.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="resolverType">
        /// The type of custom header and cell resolver that derived from <see cref="IColumnResolver{TTarget}"/>.
        /// </param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, Expression<Func<T, object>> propertySelector, Type resolverType = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            return mapper.Map(columnIndex, pi, resolverType);
        }
    }
}
