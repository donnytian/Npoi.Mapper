using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper;

namespace test.Sample
{
    /// <summary>
    /// Intend to handle all unrecognized columns.
    /// Use IsColumnMapped method to determine whether you want to map a specific column or not.
    /// 
    /// The test of IsColumnMapped will be applied to all unmapped columns if the resolver is associated to a unmapped column,
    /// which means mapped explicitly by method or attribute.
    /// 
    /// Also you can achieve this by set Mapper's DefaultResolverType.
    /// </summary>
    public class DefaultColumnResolver : IColumnResolver<SampleClass>
    {
        public bool IsColumnMapped(ref object value, int index)
        {
            try
            {
                // Custom logic to determine whether or not to map and include this column.
                // Header value is either in string or double. Try convert by needs.
                if (index > 40 && index <= 50 && value is double)
                {
                    // Assign back header value and use it from TryResolveCell method.
                    value = DateTime.FromOADate((double)value);

                    return true;
                }
            }
            catch
            {
                // Does nothing here and return false eventually.
            }

            return false;
        }

        public bool TryResolveCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            if (!(columnInfo.HeaderValue is DateTime)) return false;

            // Custom logic to handle the cell value.
            target.GeneralCollectionProperty.Add(((DateTime)columnInfo.HeaderValue).ToLongDateString() + cellValue);

            return true;
        }
    }
}
