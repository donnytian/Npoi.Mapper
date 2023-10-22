using System;
using System.Collections.Generic;

// ReSharper disable once CheckNamespace
namespace Npoi.Mapper
{
    /// <summary>
    /// Provides extensions for <see cref="Type"/> class.
    /// </summary>
    public static class TypeExtensions
    {
        /// <summary>
        /// Collection of numeric types.
        /// </summary>
        private static readonly List<Type> NumericTypes = new()
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double),
        };

        /// <summary>
        /// Check if the given type is a numeric type.
        /// </summary>
        /// <param name="type">The type to be checked.</param>
        /// <returns><c>true</c> if it's numeric; otherwise <c>false</c>.</returns>
        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type);
        }

        /// <summary>
        /// Check if the given type can be exported directly as cell value.
        /// </summary>
        /// <param name="type">The type to be checked.</param>
        public static bool CanBeExported(this Type type)
        {
            var typeToCheck = Nullable.GetUnderlyingType(type) ?? type;
            return typeToCheck.IsEnum ||
                typeToCheck == typeof(string) ||
                typeToCheck == typeof(DateTime) ||
                typeToCheck == typeof(DateTimeOffset) ||
                typeToCheck == typeof(TimeSpan) ||
                typeToCheck == typeof(Guid) ||
                NumericTypes.Contains(typeToCheck);
        }
    }
}
