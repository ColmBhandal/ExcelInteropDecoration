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
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace InteropDecoration.Decorator
{
    internal class DecoratorFactoryImpl : IDecoratorFactory
    {
        private IInteropDAPI _interopDAPI;

        public DecoratorFactoryImpl(IInteropDAPI interopDAPI)
        {
            _interopDAPI = interopDAPI;
        }

        public IApplicationD ApplicationD(Application application) => new ApplicationDImpl(_interopDAPI, application);

        public ICommentD CommentD(Comment comment) => new CommentDImpl(_interopDAPI, comment);

        public ICommentsD CommentsD(Comments comments) => new CommentsDImpl(_interopDAPI, comments);

        public IInteriorD InteriorD(Interior interior) => new InteriorDImpl(_interopDAPI, interior);

        public IListObjectsD ListObjectsD(ListObjects listObjects) => new ListObjectsDImpl(_interopDAPI, listObjects);

        public INameD NameD(Name name) => new NameDImpl(_interopDAPI, name);

        public INamesD NamesD(Names names) => new NamesDImpl(_interopDAPI, names);

        public IRangeD RangeD(Range range) => new RangeDImpl(_interopDAPI, range);

        public ITabD TabD(Tab tab) => new TabDImpl(_interopDAPI, tab);

        public IVBComponentD VBComponentD(VBComponent vbComp) => new VBComponentDImpl(_interopDAPI, vbComp);

        public IWorkbookD WorkbookD(Workbook workbook) => new WorkbookDImpl(_interopDAPI, workbook);

        public IWorkbooksD WorkbooksD(Workbooks workbooks) => new WorkbooksDImpl(_interopDAPI, workbooks);

        public IWorksheetD WorksheetD(Worksheet sheet) => new WorksheetDImpl(_interopDAPI, sheet);

        public ISheetsD WorksheetsD(Sheets worksheetsRaw) => new SheetsDImpl(_interopDAPI, worksheetsRaw);
    }
}
