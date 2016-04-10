using System;
using System.Linq;
using Npoi.Mapper;

namespace test.Sample
{
    public class MultiColumnContainerResolver : IColumnResolver<SampleClass>
    {
        public bool IsColumnMapped(ref object headerValue, int index)
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

        public bool TryTakeCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            if (!(columnInfo.HeaderValue is DateTime)) return false;

            // Custom logic to handle the cell value.
            target.CollectionGenericProperty.Add(((DateTime)columnInfo.HeaderValue).ToLongDateString() + cellValue);

            return true;
        }

        public bool TryPutCell(ColumnInfo<SampleClass> columnInfo, out object cellValue, SampleClass source)
        {
            cellValue = null;

            // Note: return false to indicate a failure; and that will increase error count.
            if (!(columnInfo?.HeaderValue is DateTime)) return false;

            var s = ((DateTime)columnInfo.HeaderValue).ToLongDateString();

            // Custom logic to set the cell value.
            if (source.CollectionGenericProperty.Count > 0 && columnInfo.Attribute.Index == 31)
            {
                cellValue = source.CollectionGenericProperty.ToList()[0].Remove(0, s.Length);
            }
            else if (source.CollectionGenericProperty.Count > 1 && columnInfo.Attribute.Index == 33)
            {
                cellValue = source.CollectionGenericProperty.ToList()[1].Remove(0, s.Length);
            }

            return true;
        }
    }
}
