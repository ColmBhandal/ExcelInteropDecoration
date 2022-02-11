using ExcelInteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ExcelInteropDecoration.Decorator.sheets
{
    public interface ISheetsD
    {
        int Count { get; }
        IEnumerable<IWorksheetD> WorksheetEnumerable();
        Sheets Worksheets { get; }
        IWorksheetD this[string index] { get; }
        IWorksheetD AddNewSheet(string sheetName);
        bool HasSheet(string sheetName);
        IWorksheetD? GetWorksheetDByNameOrNull(string sheetName);
    }
}