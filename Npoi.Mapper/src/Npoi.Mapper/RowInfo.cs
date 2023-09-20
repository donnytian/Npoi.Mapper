namespace Npoi.Mapper;

/// <summary>
/// Information for one row that read from file.
/// </summary>
/// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
public class RowInfo<TTarget>
{
    #region Properties

    /// <summary>
    /// Row number.
    /// </summary>
    public int RowNumber { get; set; }

    /// <summary>
    /// Constructed value object from the row.
    /// </summary>
    public TTarget Value { get;  set; }

    /// <summary>
    /// Column index of the first error.
    /// </summary>
    public int ErrorColumnIndex { get; set; }

    /// <summary>
    /// Error message for the first error.
    /// </summary>
    public string ErrorMessage { get; set; }

    /// <summary>
    /// Object that associated with the current row.
    /// </summary>
    public object RowTag { get; set; }

    #endregion

    #region Constructors

    /// <summary>
    /// Initialize a new RowData object.
    /// </summary>
    /// <param name="rowNumber">The row number</param>
    /// <param name="value">Constructed value object from the row</param>
    /// <param name="errorColumnIndex">Column index of the first error cell</param>
    /// <param name="errorMessage">The error message</param>
    public RowInfo(int rowNumber, TTarget value, int errorColumnIndex, string errorMessage)
    {
        RowNumber = rowNumber;
        Value = value;
        ErrorColumnIndex = errorColumnIndex;
        ErrorMessage = errorMessage;
    }

    #endregion
}
