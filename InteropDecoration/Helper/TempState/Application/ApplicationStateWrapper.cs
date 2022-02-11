using InteropDecoration.Decorator.application;
using InteropDecoration.Helper.TempState.ComAddin;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Helper.TempState.Application
{
    internal class ApplicationStateWrapper : IApplicationState
    {
        private IApplicationD Application { get; }

        public bool? CalculateBeforeSaveOrNull => Application.CalculateBeforeSaveOrNull;

        public bool? VisibleOrNull => Application.VisibleOrNull;

        public bool CalculateBeforeSave { set => Application.CalculateBeforeSave = value; }
        public bool Visible { set => Application.Visible = value; }
        public XlCalculationInterruptKey CalculationInterruptKey { get => Application.CalculationInterruptKey; set => Application.CalculationInterruptKey = value; }
        public string StatusBar { get => Application.StatusBar; set => Application.StatusBar = value; }
        public IDictionary<string, IComAddinState> ComAddinStates { get => Application.ComAddinStates; set => Application.ComAddinStates = value; }

        public ApplicationStateWrapper(IApplicationD application)
        {
            Application = application ?? throw new ArgumentNullException(nameof(application));
        }
    }
}
