namespace Npoi.Mapper
{
    /// <summary>
    /// Base contract for column resolver.
    /// Implement this interface to resolve header and cells for column(s).
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public interface IColumnResolver<TTarget>
    {
        /// <summary>
        /// Determine whether the given column mapped by this resolver or not.
        /// </summary>
        /// <param name="headerValue">
        /// Header value that is either string or double.
        /// Resolved value can be assigned back so that it can be used as HeaderValue property
        /// of columnInfo in TryResolveCell method.
        /// </param>
        /// <param name="index">Column index</param>
        /// <returns>True if can map header and column; otherwise false.</returns>
        bool IsColumnMapped(ref object headerValue, int index);

        /// <summary>
        /// Try take cell value for the given column when import data from file.
        /// </summary>
        /// <param name="columnInfo">The column info.</param>
        /// <param name="cellValue">The cell value object that is either string or double.</param>
        /// <param name="target">The target object of the mapped type.</param>
        /// <returns>True if cell was resolved without error; otherwise false.</returns>
        bool TryTakeCell(ColumnInfo<TTarget> columnInfo, object cellValue, TTarget target);

        /// <summary>
        /// Try set value to cell for the given column when export object to file.
        /// </summary>
        /// <param name="columnInfo">The column info.</param>
        /// <param name="cellValue">The value that will be set to cell.</param>
        /// <param name="source">The object of the mapped type.</param>
        /// <returns>True if cell was resolved without error; otherwise false.</returns>
        bool TryPutCell(ColumnInfo<TTarget> columnInfo, out object cellValue, TTarget source);
    }
}
