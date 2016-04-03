using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
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
        public int Index { get; internal set; } = -1;

        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; internal set; }

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
        internal bool? UseLastNonBlankValue { get; set; }

        /// <summary>
        /// Indicate whether to ignore the property.
        /// </summary>
        internal bool? Ignored { get; set; }

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

        #region Constructors

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnAttribute"/> class.
        /// </summary>
        public ColumnAttribute()
        {
        }

        /// <summary>
        /// Initialize a new instance of <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        public ColumnAttribute(ushort index)
        {
            Index = index;
        }
        /// <summary>
        /// Initialize a new instance of <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        public ColumnAttribute(string name)
        {
            Name = name;
        }

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

        /// <summary>
        /// Merge properties from a source <see cref="ColumnAttribute"/> object.
        /// All properties will be updated from source's specified properties.
        /// </summary>
        /// <param name="source">The object to merge from.</param>
        /// <param name="overwrite">
        /// Whether or not to overwrite specified properties from source if source's properties are specified.
        /// Note that Index and Name are considered together as one key property.
        /// </param>
        public void MergeFrom(ColumnAttribute source, bool overwrite = true)
        {
            if (source == null) return;

            if (source.Index >= 0 || source.Name != null)
            {
                if (overwrite || (Index < 0 && Name == null))
                {
                    Index = source.Index;
                    Name = source.Name;
                }
            }

            if (source.Property != null && (overwrite || Property == null)) Property = source.Property;
            if (source.ResolverType != null && (overwrite || ResolverType == null)) ResolverType = source.ResolverType;
            if (source.UseLastNonBlankValue != null && (overwrite || UseLastNonBlankValue == null)) UseLastNonBlankValue = source.UseLastNonBlankValue;
            if (source.Ignored != null && (overwrite || Ignored == null)) Ignored = source.Ignored;
            if (source.BuiltinFormat != 0 && (overwrite || BuiltinFormat == 0)) BuiltinFormat = source.BuiltinFormat;
            if (source.CustomFormat != null && (overwrite || CustomFormat == null)) CustomFormat = source.CustomFormat;
            // TODO: fix bug for Mapper.Format(0) and Mapper.Format(null);
        }

        /// <summary>
        /// Merge properties to a attribute dictionary.
        /// </summary>
        /// <param name="attributes">The dictionary to be merged into.</param>
        /// <param name="overwrite">
        /// Whether or not to overwrite specified properties to existed object if that object's properties are specified.
        /// Note that Index and Name are considered together as one key property.
        /// </param>
        public void MergeTo(Dictionary<PropertyInfo, ColumnAttribute> attributes, bool overwrite = true)
        {
            if (attributes == null) return;
            var pi = Property;
            if (pi == null) return;

            var existed = attributes.ContainsKey(pi) ? attributes[pi] : null;
            var isIndexSet = Index >= 0;

            if (isIndexSet && !overwrite)
                if (attributes.Any(p => p.Key != pi && p.Value.Index == Index))
                {
                    // Clear Index if there is same index already set (with overwrite = false).
                    Index = -1;
                    isIndexSet = false;
                }

            if (existed != null)
            {
                isIndexSet = isIndexSet && ((existed.Index != Index) || overwrite);
                existed.MergeFrom(this, overwrite);
                isIndexSet = isIndexSet && (existed.Index == Index);
            }
            else
            {
                attributes[pi] = this;
            }

            if (isIndexSet) // True if the index set successfully, otherwise it's been ignored/ cleared.
            {
                // Clear other attributes' Index if they have same index.
                attributes.Where(p => p.Key != pi && p.Value.Index == Index).ForEach(p => p.Value.Index = -1);
            }
        }

        #endregion
    }
}
