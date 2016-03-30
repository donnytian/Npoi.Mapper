using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npoi.Mapper;

namespace test.Sample
{
    /// <summary>
    /// Sample class to test <see cref="ColumnResolver{TTarget}"/> class.
    /// </summary>
    public class SampleColumnResolver : ColumnResolver<SampleClass>
    {
        public override bool TryResolveCell(ColumnInfo<SampleClass> columnInfo, object cellValue, SampleClass target)
        {
            throw new NotImplementedException();
        }

        public override bool TryResolveHeader(ref object value, int index)
        {
            throw new NotImplementedException();
        }
    }
}
