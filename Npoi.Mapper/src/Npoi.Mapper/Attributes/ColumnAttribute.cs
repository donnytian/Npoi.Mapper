using System;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace Npoi.Mapper.Attributes
{
    /// <summary>
    /// Specifies attributes for a property that is going to map to a column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Global")]
    [SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Local")]
    public sealed class ColumnAttribute : Attribute
    {
        #region Properties

        /// <summary>
        /// Column index.
        /// </summary>
        public int Index { get; set; } = -1;

        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Mapped property for this column.
        /// </summary>
        public PropertyInfo Property { get; internal set; }

        /// <summary>
        /// The type of class that is derived from <see cref="ColumnResolver{TTarget}"/> class.
        /// </summary>
        public Type ResolverType { get; set; }

        /// <summary>
        /// Indicate whether to use the last non-blank value.
        /// Typically handle the blank error in merged cells.
        /// </summary>
        public bool UseLastNonBlankValue { get; set; }

        /// <summary>
        /// Indicate whether to ignore the property.
        /// </summary>
        public bool Ignored { get; set; }

        /// <summary>
        /// Gets or sets the built-in format, see https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html for possible values.
        /// </summary>
        /// <value>
        /// The built-in format.
        /// </value>
        public short BuiltinFormat { get; set; }

        /// <summary>
        /// Gets or sets the custom format, see https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 for the syntax.
        /// </summary>
        /// <value>
        /// The custom format.
        /// </value>
        public string CustomFormat { get; set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Get a member wise clone of this object.
        /// </summary>
        /// <returns></returns>
        public ColumnAttribute Clone()
        {
            return (ColumnAttribute)MemberwiseClone();
        }

        #endregion
    }
}
