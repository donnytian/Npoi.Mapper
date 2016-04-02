using System;
using Npoi.Mapper;

namespace test.Sample
{
    public class MultiColumnContainerResolver : ColumnResolver<SampleClass>
    {
        public override bool IsColumnMapped(ref object headerValue, int index)
        {
            try
            {
                // Custom logic to determine whether or not to map and include this column.
                // Header value is either in string or double. Try convert by needs.
                if (index > 30 && index <= 40 && headerValue is double)
                {
                    // Assign back header value and use it from TryResolveCell method.
                    headerValue = DateTime.FromOADate((double)headerValue);

                    return true;
                }
            }
            catch
            {
                // Does nothing here and return false eventually.
            }

            return false;
        }

        public override bool TryResolveCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            if (!(columnInfo.HeaderValue is DateTime)) return false;

            // Custom logic to handle the cell value.
            target.CollectionGenericProperty.Add(((DateTime)columnInfo.HeaderValue).ToLongDateString() + cellValue);

            return true;
        }
    }
}
