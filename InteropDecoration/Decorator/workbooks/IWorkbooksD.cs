using Microsoft.Office.Interop.Excel;
using CsharpExtras.Enumerable.NonEmpty;
using InteropDecoration.Decorator.workbook;
using InteropDecoration._base;

namespace InteropDecoration.Decorator.workbooks
{
    public interface IWorkbooksD
    {
        Workbooks RawWorkbooks { get; }

        IWorkbookD AddWorkbookWithSheets(INonEmptyEnumerable<string> sheetName);

        IWorkbookD Open(string filePath);
    }

    public interface IWorkbooksDBuilder : IBuilder<IWorkbooksD>
    {
        IWorkbooksDBuilder WithWorkbooks(Workbooks obj);
    }

    public interface IWorkbooksDBuilderFactory : IFactory<IWorkbooksDBuilder>
    {
    }
}
