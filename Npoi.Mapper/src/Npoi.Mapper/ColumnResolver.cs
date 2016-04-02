namespace Npoi.Mapper
{
    /// <summary>
    /// Use derived of this class to resolve header and cells for a column.
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public abstract class ColumnResolver<TTarget>
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
        public abstract bool IsColumnMapped(ref object headerValue, int index);

        /// <summary>
        /// Try resolve cell.
        /// </summary>
        /// <param name="columnInfo">The column info.</param>
        /// <param name="cellValue">The cell value object that is either string or double.</param>
        /// <param name="target">The target object of the mapped type.</param>
        /// <returns>True if cell was resolved without error; otherwise false.</returns>
        public abstract bool TryResolveCell(ColumnInfo<TTarget> columnInfo, object cellValue, TTarget target);
    }
}
