namespace Npoi.Mapper
{
    /// <summary>
    /// Use derived of this class to resolve header and cells for a column.
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public abstract class ColumnResolver<TTarget>
    {
        /// <summary>
        /// Try resolve header.
        /// </summary>
        /// <param name="value">
        /// Header value that is either string or double.
        /// Resolved value can be assigned back and will be passed in again as HeaderValue property
        /// of columnInfo in TryResolveCell method.
        /// </param>
        /// <param name="index">Column index</param>
        /// <returns>True if can take header and column to match property; otherwise false.</returns>
        public abstract bool TryResolveHeader(ref object value, int index);

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
