using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelInteropDecoration.Decorator.comment
{
    public interface ICommentD
    {
        [Obsolete]
        Comment RawComment { get; }
        IRangeD ParentCell { get; }

        string Text { get; set; }
    }
}