using ExcelInteropDecoration.Decorator._base;
using ExcelInteropDecoration.Decorator.comment;

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Decorator.comments
{
    class CommentsDImpl : DecoratorBase, ICommentsD
    {
        public Comments RawComments { get; }

        public CommentsDImpl(IInteropDAPI api, Comments rawComments) : base(api)
        {
            RawComments = rawComments ?? throw new ArgumentNullException(nameof(rawComments));
        }

        public int Count => RawComments.Count;
        public ICommentD Item(int index) => DecoratorFactory.CommentD(RawComments.Item(index));

        public ISet<ICommentD> AsSet()
        {
            ISet<ICommentD> set = new HashSet<ICommentD>();
            foreach(ICommentD comment in this)
            {
                set.Add(comment);
            }            
            return set;
        }

        public IEnumerator<ICommentD> GetEnumerator()
        {
            foreach(object rawComment in RawComments)
            {
                if (rawComment is Comment comment)
                {
                    yield return DecoratorFactory.CommentD(comment);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}