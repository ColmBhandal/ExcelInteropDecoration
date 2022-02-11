using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelInteropDecoration.Decorator.listObjects
{
    public interface IListObjectD
    {
        ListObject RawListObject { get; }
        IRangeD Range { get; }
        IRangeD HeaderRowRange { get; }
        IRangeD DataBodyRange { get; }

        void ProcessEachTableDataRow(Action<int, Func<int, string>> rowDataProcessor);
    }
}