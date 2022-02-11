using InteropDecoration.Decorator.application;
using InteropDecoration.Decorator.comment;
using InteropDecoration.Decorator.comments;
using InteropDecoration.Decorator.interior;
using InteropDecoration.Decorator.listObjects;
using InteropDecoration.Decorator.names;
using InteropDecoration.Decorator.range;
using InteropDecoration.Decorator.sheets;
using InteropDecoration.Decorator.tab;
using InteropDecoration.Decorator.vbComponent;
using InteropDecoration.Decorator.workbook;
using InteropDecoration.Decorator.workbooks;
using InteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace InteropDecoration.Decorator
{
    public interface IDecoratorFactory
    {
        IWorkbookD WorkbookD(Workbook workbook);
        IRangeD RangeD(Range range);
        ICommentD CommentD(Comment comment);
        IWorksheetD WorksheetD(Worksheet sheet);
        IInteriorD InteriorD(Interior interior);
        INameD NameD(Name name);
        ISheetsD WorksheetsD(Sheets worksheetsRaw);
        IApplicationD ApplicationD(Application application);
        IVBComponentD VBComponentD(Microsoft.Vbe.Interop.VBComponent vbComp);
        INamesD NamesD(Names names);
        ICommentsD CommentsD(Comments comments);
        ITabD TabD(Tab tab);
        IListObjectsD ListObjectsD(ListObjects listObjects);
        IWorkbooksD WorkbooksD(Workbooks workbooks);
    }
}