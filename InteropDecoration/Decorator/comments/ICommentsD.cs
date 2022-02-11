using InteropDecoration.Decorator.comment;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace InteropDecoration.Decorator.comments
{
    public interface ICommentsD : IEnumerable<ICommentD>
    {
        int Count { get; }
        Comments RawComments { get; }

        ISet<ICommentD> AsSet();
        ICommentD Item(int index);
    }
}