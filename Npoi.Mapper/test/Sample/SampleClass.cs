using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper.Attributes;

namespace test.Sample
{
    /// <summary>
    /// Sample class for testing purpose.
    /// </summary>
    public class SampleClass
    {
        public SampleClass()
        {
            CollectionGenericProperty = new List<string>();
        }

        public SampleClass(ICollection<string> collectionGenericProperty)
        {
            CollectionGenericProperty = collectionGenericProperty;
        }

        public string StringProperty { get; set; }

        public int Int32Property { get; set; }

        public DateTime DateProperty { get; set; }

        public double DoubleProperty { get; set; }

        public SampleEnum EnumProperty { get; set; }

        public object ObjectProperty { get; set; }

        public ICollection<string> CollectionGenericProperty { get; set; }

        [ColumnName("By Name")]
        public string ColumnNameAttributeProperty { get; set; }

        [ColumnIndex(1)]
        public string ColumnIndexAttributeProperty { get; set; }

        [UseLastNonBlankValue]
        public string UserLastNonBlankValueAttributeProperty { get; set; }
    }
}
