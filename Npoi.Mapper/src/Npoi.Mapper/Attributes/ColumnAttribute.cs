using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        /// Property name, this is only used for dynamic type.
        /// </summary>
        public string PropertyName { get; set; }

        private PropertyInfo _property;
        /// <summary>
        /// Mapped property for this column.
        /// </summary>
        public PropertyInfo Property
        {
            get => _property;

            internal set
            {
                _property = value;

                if (value != null)
                {
                    PropertyUnderlyingType = Nullable.GetUnderlyingType(value.PropertyType);
                    PropertyUnderlyingConverter = PropertyUnderlyingType != null ? TypeDescriptor.GetConverter(PropertyUnderlyingType) : null;
                }
                else
                {
                    PropertyUnderlyingType = null;
                    PropertyUnderlyingConverter = null;
                }
            }
        }

        /// <summary>
        /// Get underlying type if property is nullable value type, otherwise return null.
        /// </summary>
        public Type PropertyUnderlyingType { get; private set; }

        /// <summary>
        /// Get converter if property is nullable value type.
        /// </summary>
        public TypeConverter PropertyUnderlyingConverter { get; private set; }

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
        /// Gets or sets the custom format, see https://support.office.com/en-us/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 for the syntax.
        /// </summary>
        public string CustomFormat { get; set; }

        /// <summary>
        /// Indicates whether or not to ignore all errors for the column.
        /// </summary>
        public bool? IgnoreErrors { get; set; }

        /// <summary>
        /// Try take cell value for the given column when import data from file.
        /// </summary>
        internal Func<IColumnInfo, object, bool> TryTake { get; set; }

        /// <summary>
        /// Try set value to cell for the given column when export object to file.
        /// </summary>
        internal Func<IColumnInfo, object, bool> TryPut { get; set; }

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
        /// <returns>The member wise clone.</returns>
        public ColumnAttribute Clone()
        {
            return (ColumnAttribute)MemberwiseClone();
        }

        /// <summary>
        /// Get a member wise clone of this object with given index.
        /// </summary>
        /// <param name="index">The index of column.</param>
        /// <returns>The member wise clone with specified index.</returns>
        public ColumnAttribute Clone(int index)
        {
            var clone = Clone();
            clone.Index = index;
            return clone;
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
            if (source.UseLastNonBlankValue != null && (overwrite || UseLastNonBlankValue == null)) UseLastNonBlankValue = source.UseLastNonBlankValue;
            if (source.Ignored != null && (overwrite || Ignored == null)) Ignored = source.Ignored;
            if (source.CustomFormat != null && (overwrite || CustomFormat == null)) CustomFormat = source.CustomFormat;
            if (source.IgnoreErrors != null && (overwrite || IgnoreErrors == null)) IgnoreErrors = source.IgnoreErrors;
            // TODO: fix for Mapper.Format(0) and Mapper.Format(null);

            if (overwrite || TryPut == null) TryPut = source.TryPut;
            if (overwrite || TryTake == null) TryTake = source.TryTake;
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
