using ExcelInteropDecoration.Decorator.application;
using ExcelInteropDecoration.Decorator.comment;
using ExcelInteropDecoration.Decorator.comments;
using ExcelInteropDecoration.Decorator.interior;
using ExcelInteropDecoration.Decorator.listObjects;
using ExcelInteropDecoration.Decorator.names;
using ExcelInteropDecoration.Decorator.range;
using ExcelInteropDecoration.Decorator.sheets;
using ExcelInteropDecoration.Decorator.tab;
using ExcelInteropDecoration.Decorator.vbComponent;
using ExcelInteropDecoration.Decorator.workbook;
using ExcelInteropDecoration.Decorator.workbooks;
using ExcelInteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelInteropDecoration.Decorator
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