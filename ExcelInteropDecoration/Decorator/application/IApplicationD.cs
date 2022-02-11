using ExcelInteropDecoration.Decorator.range;
using ExcelInteropDecoration.Decorator.workbooks;
using ExcelInteropDecoration.Decorator.worksheet;
using ExcelInteropDecoration.Helper.TempState.Application;
using ExcelInteropDecoration.Helper.TempState.ComAddin;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelInteropDecoration.Decorator.application
{
    public interface IApplicationD
    {
        Application RawApplication { get; }

        IRangeD Selection { get; }
        //Non-mvp: add an active sheet that returns a worksheet type vs. a decorator type
        IWorksheetD ActiveSheetD { get; }
        IRangeD ActiveCell { get; }

        IWorkbooksD Workbooks { get; }
        Window ActiveWindow { get; }

        bool ScreenUpdating { get; set; }
        bool DisplayAlerts { get; set; }
        bool? VisibleOrNull { get; }
        bool? CalculateBeforeSaveOrNull { get; }
        bool CalculateBeforeSave { set; }
        bool Visible { set; }
        XlCalculationInterruptKey CalculationInterruptKey { get; set; }
        string StatusBar { get; set; }
        IDictionary<string, IComAddinState> ComAddinStates { get; set; }

        void RunWithScreenUpdateOff(System.Action doWork);
        void RunWithDisplayAlertsOff(System.Action actionToRun);
        void QuitWithoutPrompt();
        IApplicationTempState NewTempState();
        void RunWithStatusBarMessage(System.Action action, string message);
        void UnbreakableCalculate();
        bool TryCompileVbaProject();
        IRangeD Range(string address);
        IRangeD Intersect(IRangeD range1, IRangeD range2);

        void ClearUndoList();
        void ClearCutCopyMode();
        IRangeD Union(IRangeD range1, IRangeD range2);
    }
}