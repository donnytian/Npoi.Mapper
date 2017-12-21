using System;
using System.Linq.Expressions;
using System.Reflection;
using Npoi.Mapper.Attributes;

namespace Npoi.Mapper
{
    /// <summary>
    /// Extension methods for mapping.
    /// </summary>
    public static class MapExtensions
    {
        /// <summary>
        /// Map property to a column by specified column name and <see cref="PropertyInfo"/>.
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnName">The column name.</param>
        /// <param name="propertyInfo">The <see cref="PropertyInfo"/> object.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map(this Mapper mapper, string columnName, PropertyInfo propertyInfo,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));

            var columnAttribute = new ColumnAttribute
            {
                Property = propertyInfo,
                Name = columnName,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based).
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertyInfo">The <see cref="PropertyInfo"/> object.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map(this Mapper mapper, ushort columnIndex, PropertyInfo propertyInfo,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));

            var columnAttribute = new ColumnAttribute
            {
                Property = propertyInfo,
                Index = columnIndex,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
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

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property selector.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <param name="exportedColumnName">The column name for export.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, Expression<Func<T, object>> propertySelector, string exportedColumnName,
            Func<IColumnInfo, object, bool> tryTake = null,
            Func<IColumnInfo, object, bool> tryPut = null)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            var columnAttribute = new ColumnAttribute
            {
                Property = pi,
                Index = columnIndex,
                Name = exportedColumnName,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Map property to a column by specified column index(zero-based) and property name.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <param name="propertyName">The property name.</param>
        /// <param name="exportedColumnName">The column name for export.</param>
        /// <param name="tryTake">The function try to import from cell value to the target object.</param>
        /// <param name="tryPut">The function try to export source object to the cell.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Map<T>(this Mapper mapper, ushort columnIndex, string propertyName, string exportedColumnName,
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
                Name = exportedColumnName,
                TryPut = tryPut,
                TryTake = tryTake,
                Ignored = false
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Ignores all errors for the specified property.
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="propertyName">The property name.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper IgnoreErrorsFor<T>(this Mapper mapper, string propertyName)
        {
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));

            var type = typeof(T);
            var pi = type.GetProperty(propertyName, MapHelper.BindingFlag);

            if (pi == null && type != typeof(object)) throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");

            var columnAttribute = new ColumnAttribute
            {
                Property = pi,
                PropertyName = propertyName,
                IgnoreErrors = true
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Ignores all errors for the specified property.
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="propertySelector">Property selector.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper IgnoreErrorsFor<T>(this Mapper mapper, Expression<Func<T, object>> propertySelector)
        {
            var pi = MapHelper.GetPropertyInfoByExpression(propertySelector);
            if (pi == null) throw new InvalidOperationException($"Cannot find the property specified by the selector.");

            var columnAttribute = new ColumnAttribute
            {
                Property = pi,
                IgnoreErrors = true
            };

            return mapper.Map(columnAttribute);
        }

        /// <summary>
        /// Ignore property by names. Ignored properties will not be mapped for import and export.
        /// </summary>
        /// <typeparam name="T">The target object type.</typeparam>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="propertyNames">Property names.</param>
        /// <returns>The mapper object.</returns>
        public static Mapper Ignore<T>(this Mapper mapper, params string[] propertyNames)
        {
            var type = typeof(T);

            foreach (var propertyName in propertyNames)
            {
                var pi = type.GetProperty(propertyName, MapHelper.BindingFlag);

                if (pi == null && type != typeof(object)) // Does not throw for dynamic type.
                {
                    throw new InvalidOperationException($"Cannot find a public property in name of '{propertyName}'.");
                }

                var columnAttribute = new ColumnAttribute
                {
                    Property = pi,
                    PropertyName = propertyName,
                    Ignored = true
                };
                mapper.Map(columnAttribute);
            }

            return mapper;
        }
    }
}
