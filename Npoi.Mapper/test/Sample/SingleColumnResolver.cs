using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper;

namespace test.Sample
{
    /// <summary>
    /// Intend to handle only mapped column.
    /// 
    /// The resolver will be only applied to one column if the resolver is associated to a mapped column explicitly
    /// by method or attribute.
    /// 
    /// Return value of IsColumnMapped method will be ignored for a single column resolver, 
    /// But you can still use this method to change header value in order to use it in TryResolveCell method.
    /// Header value is either in string or double (even for int and date type). Try convert by needs.
    /// </summary>
    public class SingleColumnResolver : IColumnResolver<SampleClass>
    {
        public bool IsColumnMapped(ref object value, int index)
        {
            try
            {
                // Return value of the method will be ignored. But you can change header value.
                // Because in Excel, cell value is either in string or double. Try convert by needs.
                if (index > 50 && index <= 60 && value is double)
                {
                    // Assign back header value and use it from TryResolveCell method.
                    value = DateTime.FromOADate((double)value);
                }
            }
            catch
            {
                // Does nothing here.
            }

            return true;
        }

        public bool TryTakeCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            if (!(columnInfo.HeaderValue is DateTime)) return false;

            // Custom logic to get the cell value.
            target.SingleColumnResolverProperty =((DateTime)columnInfo.HeaderValue).ToLongDateString() + cellValue;

            return true;
        }

        public bool TryPutCell(ColumnInfo<SampleClass> columnInfo, out object cellValue, SampleClass source)
        {
            cellValue = null;

            // Note: return false to indicate a failure; and that will increase error count.
            if (!(columnInfo?.HeaderValue is DateTime)) return false;

            var s = ((DateTime)columnInfo.HeaderValue).ToLongDateString();

            // Custom logic to set the cell value.
            cellValue = source.SingleColumnResolverProperty?.Remove(0, s.Length);

            return true;
        }
    }
}
