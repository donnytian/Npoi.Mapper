using Npoi.Mapper.Attributes;

namespace Npoi.Mapper
{
    /// <summary>
    /// Information required for one column when mapping between object and file rows.
    /// </summary>
    public interface IColumnInfo
    {
        /// <summary>
        /// Value for the column header.
        /// </summary>
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        object HeaderValue { get; }

        /// <summary>
        /// The mapped property information.
        /// </summary>
        ColumnAttribute Attribute { get; }

        /// <summary>
        /// The last non-blank value.
        /// </summary>
        object LastNonBlankValue { get; set; }

        /// <summary>
        /// Get the header cell format.
        /// </summary>
        short? HeaderFormat { get; set; }

        /// <summary>
        /// Get the data cell format.
        /// </summary>
        short? DataFormat { get; set; }
    }
}
