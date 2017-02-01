using System;
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
        object HeaderValue { get; set; }

        /// <summary>
        /// The mapped property information.
        /// </summary>
        ColumnAttribute Attribute { get; }

        /// <summary>
        /// The last non-blank cell value.
        /// </summary>
        object LastNonBlankValue { get; set; }

        /// <summary>
        /// The current cell value, might be used for custom resolving.
        /// </summary>
        object CurrentValue { get; set; }

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
