using InteropDecoration.Helper.TempState.ComAddin;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Helper.TempState.Application
{
    internal class ApplicationStatePoco : IApplicationState
    {
        public bool? CalculateBeforeSaveOrNull { get; private set; }
        public bool? VisibleOrNull { get; private set; }

        public XlCalculationInterruptKey CalculationInterruptKey { get; set; }
        public string StatusBar { get; set; } = "";
        public IDictionary<string, IComAddinState> ComAddinStates { get; set; }
            = new Dictionary<string, IComAddinState>();
        public bool CalculateBeforeSave
        {
            set
            {
                CalculateBeforeSaveOrNull = value;
            }
        }
        public bool Visible
        {
            set
            {
                VisibleOrNull = value;
            }
        }
    }
}
