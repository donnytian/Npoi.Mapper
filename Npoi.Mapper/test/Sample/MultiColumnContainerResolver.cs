using System;
using Npoi.Mapper;

namespace test.Sample
{
    public class MultiColumnContainerResolver : ColumnResolver<SampleClass>
    {
        public override bool TryResolveHeader(ref object value, int index)
        {
            try
            {
                // Custom logic to determin whether or not to take this column.
                if (index > 5 && value is double)
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

        public override bool TryResolveCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            // Note: return false to indicate a failure; and that will increase error count.
            if (columnInfo?.HeaderValue == null || cellValue == null) return false;

            target.CollectionGenericProperty.Add(columnInfo.HeaderValue.ToString() + cellValue);

            return true;
        }
    }
}
