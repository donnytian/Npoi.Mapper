using System;
using System.Collections.Generic;

namespace Npoi.Mapper
{
    /// <summary>
    /// Provides extensions for <see cref="IEnumerable{T}"/> class.
    /// </summary>
    public static class EnumerableExtensions
    {
        /// <summary>
        /// Extension for <see cref="IEnumerable{T}"/> object to handle each item.
        /// </summary>
        /// <typeparam name="T">The item type.</typeparam>
        /// <param name="sequence">The enumerable sequence.</param>
        /// <param name="action">Action to apply to each item.</param>
        public static void ForEach<T>(this IEnumerable<T> sequence, Action<T> action)
        {
            if (sequence == null) return;

            foreach (var item in sequence)
            {
                action(item);
            }
        }
    }
}
