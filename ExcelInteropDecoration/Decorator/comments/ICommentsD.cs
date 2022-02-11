using ExcelInteropDecoration.Decorator.comment;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ExcelInteropDecoration.Decorator.comments
{
    public interface ICommentsD : IEnumerable<ICommentD>
    {
        int Count { get; }
        Comments RawComments { get; }

        ISet<ICommentD> AsSet();
        ICommentD Item(int index);
    }
}