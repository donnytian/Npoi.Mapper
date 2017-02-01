using System;
using System.Linq.Expressions;
using Npoi.Mapper.Attributes;

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
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, string columnName, string propertyName,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));

            var type = typeof(T);
            var pi = type.GetProperty(propertyName, MapHelper.BindingFlag);

            if (pi == null && type != typeof(object)) throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");

            var columnAttribute = new ColumnAttribute
            {
                Property = pi,
                PropertyName = propertyName,
                Name = columnName,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Map property to a column by specified column name and property selector.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, string columnName, Expression<Func<T, object>> propertySelector,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));

            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);

            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            return mapper.Map(columnName, pi, tryTake, tryPut);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertyName">The property name.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, string propertyName,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));

            var type = typeof(T);
            var pi = type.GetProperty(propertyName, MapHelper.BindingFlag);

            if (pi == null && type != typeof(object)) throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");

            var columnAttribute = new ColumnAttribute
            {
                Property = pi,
                PropertyName = propertyName,
                Index = columnIndex,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property selector.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, Expression<Func<T, object>> propertySelector,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            return mapper.Map(columnIndex, pi, tryTake, tryPut);
        }
    }
}
