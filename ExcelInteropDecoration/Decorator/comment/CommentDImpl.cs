using ExcelInteropDecoration.Decorator._base;

using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelInteropDecoration.Decorator.comment
{
    class CommentDImpl : DecoratorBase, ICommentD
    {
        public Comment RawComment { get; }

        public CommentDImpl(IInteropDAPI api, Comment rawComment) : base(api)
        {
            RawComment = rawComment ?? throw new ArgumentNullException(nameof(rawComment));            
        }

        public IRangeD ParentCell => GetParentCell();

        private IRangeD GetParentCell() => InteropTypeValidator.GetMapValidate<Range, IRangeD>
            (() => RawComment.Parent, DecoratorFactory.RangeD);

        public string Text
        {
            get => RawComment.Text();
            set => SetCommentText(value);
        }

        private void SetCommentText(string value)
        {
            if (Text != value)
            {
                ParentCell.ClearComments();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    ParentCell.AddComment(value);
                }
            }
        }
    }
}
