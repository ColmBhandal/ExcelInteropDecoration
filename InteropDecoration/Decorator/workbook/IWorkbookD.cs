using CsharpExtras.Visitor;
using InteropDecoration.Decorator.application;

using InteropDecoration.Decorator.names;
using InteropDecoration.Decorator.range;
using InteropDecoration.Decorator.sheets;
using InteropDecoration.Decorator.vbComponent;
using InteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace InteropDecoration.Decorator.workbook
{
    public interface IWorkbookD
    {
        Workbook Workbook { get; }
        IWorksheetD ActiveSheet();

        IApplicationD Application { get; }

        ISheetsD Worksheets { get; }

        IEnumerable<IWorksheetD> WorksheetEnumerable();

        IEnumerable<IWorksheetD> VisibleWorksheetsEnumerable();

        string Path { get; }

        string NameNoExtension { get; }

        string Extension { get; }
        
        IRangeD NamedRange(string rangeName);
        IRangeD? NamedRangeOrNull(string rangeName);

        IWorksheetD? GetWorksheetDByNameOrNull(string sheetName);

        IVBComponentD GetVbComponentByName(string vbCompName);        

        void ImportVbModule(string filePath);

        void DeleteVbModule(IVBComponentD vBComponent);

        void Save();
        void Close();
        void CloseWithoutPrompt();
        void SaveAs(string v, XlFileFormat format = XlFileFormat.xlOpenXMLWorkbook);
        /// <summary>
        /// Saves the existing workbook to the same location with the new name given. The file format stays the same.
        /// </summary>        
        void SaveAsHere(string newName);
        void SaveAndClose();
        
        void SaveAndCloseWithoutRecalc();

        void ForceRecalculate();

        void RecalculateIfNecessary();

        bool IsRecalcPending();
        

        /// <param name="sheetCellReference">This is the reference to the range as it would appear from another sheet in the same workbook.</param>        
        IRangeD Range(string sheetCellReference);
        INamesD Names { get; }

        string Name { get; }

        ISet<string> NamesAsStrings();

        ISet<string> NamesWithErrorsAsStrings();
    }
}
