using ExcelInteropDecoration.Helper.TempState.ComAddin;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Helper.TempState.Application
{
    public interface IApplicationState
    {
        bool? CalculateBeforeSaveOrNull { get; }
        bool? VisibleOrNull { get; }
        bool CalculateBeforeSave { set; }
        bool Visible { set; }

        XlCalculationInterruptKey CalculationInterruptKey { get; set; }
        string StatusBar { get; set; }

        IDictionary<string, IComAddinState> ComAddinStates { get; set; }
    }
}
