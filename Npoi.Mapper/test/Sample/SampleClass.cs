using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper.Attributes;

namespace test.Sample
{
    /// <summary>
    /// Sample class for testing purpose.
    /// </summary>
    public class SampleClass : BaseClass
    {
        public SampleClass()
        {
            CollectionGenericProperty = new List<string>();
            GeneralCollectionProperty = new List<string>();
        }

        public SampleClass(ICollection<string> collectionGenericProperty)
        {
            CollectionGenericProperty = collectionGenericProperty;
        }

        public string StringProperty { get; set; }

        public int Int32Property { get; set; }

        public bool BoolProperty { get; set; }

        public DateTime DateProperty { get; set; }

        public double DoubleProperty { get; set; }

        public SampleEnum EnumProperty { get; set; }

        public object ObjectProperty { get; set; }

        public ICollection<string> CollectionGenericProperty { get; set; }

        public string SingleColumnResolverProperty { get; set; }

        [Column("By Name")]
        public string ColumnNameAttributeProperty { get; set; }

        [Column(11)]
        public string ColumnIndexAttributeProperty { get; set; }

        public string IndexOverNameAttributeProperty { get; set; }

        [UseLastNonBlankValue]
        public string UseLastNonBlankValueAttributeProperty { get; set; }

        [Ignore]
        public string IgnoredAttributeProperty { get; set; }

        [Display(Name = "Display Name")]
        public string DisplayNameProperty { get; set; }

        public string GeneralProperty { get; set; }

        public ICollection<string> GeneralCollectionProperty { get; set; }

        [Column(CustomFormat = "0%")]
        public double CustomFormatProperty { get; set; }
    }
}
