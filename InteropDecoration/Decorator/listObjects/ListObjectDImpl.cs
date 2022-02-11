using InteropDecoration.Decorator._base;

using InteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsharpExtras.Enumerable.OneBased;

namespace InteropDecoration.Decorator.listObjects
{
    class ListObjectDImpl : DecoratorBase, IListObjectD
    {
        public ListObjectDImpl(IInteropDAPI api, ListObject listObject) : base(api)
        {
            RawListObject = listObject;
        }

        public ListObject RawListObject { get; }

        public IRangeD Range => DecoratorFactory.RangeD(RawListObject.Range);

        public IRangeD HeaderRowRange => DecoratorFactory.RangeD(RawListObject.HeaderRowRange);

        public IRangeD DataBodyRange => DecoratorFactory.RangeD(RawListObject.DataBodyRange);

        public void ProcessEachTableDataRow(Action<int, Func<int, string>> rowDataProcessor)
        {
            IRangeD range = DataBodyRange;
            IOneBasedArray2D<string> values = range.Value;
            int lastUsedRow = values.LastUsedRow();
            for (int rowIndex = 1; rowIndex <= lastUsedRow; rowIndex++)
            {
                rowDataProcessor(rowIndex, col => values[rowIndex, col]);
            }
        }
    }
}
