using InteropDecoration.Decorator.range;
using InteropDecoration.Decorator.workbooks;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InteropDecoration.Decorator.worksheet;
using System.Threading;
using InteropDecoration.Decorator._base;
using InteropDecoration.Helper.TempState;
using Range = Microsoft.Office.Interop.Excel.Range;
using InteropDecoration.Helper.TempState.Application;
using System.Runtime.InteropServices;
using InteropDecoration.Helper.TempState.ComAddin;
using Microsoft.Office.Core;

namespace InteropDecoration.Decorator.application
{
    class ApplicationDImpl : DecoratorBase, IApplicationD
    {
        public IApplicationTempState NewTempState() => InteropDApi.NewApplicationTempState(this);
        public Application RawApplication { get; private set; }

        public bool? VisibleOrNull =>
            GetBoolAndCatchComException(() => RawApplication.Visible, "Get visible");

        public bool? CalculateBeforeSaveOrNull =>
            GetBoolAndCatchComException(() => RawApplication.CalculateBeforeSave, "Get calculate before save");
        public bool CalculateBeforeSave { set => RawApplication.CalculateBeforeSave = value; }
        public bool Visible { set => RawApplication.Visible = value; }

        private bool? GetBoolAndCatchComException(Func<bool> prop, string actionDescription)
        {
            try
            {
                bool rawVal = prop();
                if (rawVal is bool val)
                {
                    return val;
                }
                return null;
            }
            catch (COMException e)
            {
                Log.Debug($"COM Exception caught while trying: {actionDescription}", e);
                return null;
            }
        }

        public XlCalculationInterruptKey CalculationInterruptKey
        {
            get => RawApplication.CalculationInterruptKey;
            set => RawApplication.CalculationInterruptKey = value;
        }
        public ApplicationDImpl(IInteropDAPI api, Application application)
            : base(api)
        {            
            RawApplication = application;            
        }

        //Non-mvp: do we wrap this in a forwarder/transposer in case it's on a transposed sheet?
        public IRangeD ActiveCell => DecoratorFactory.RangeD(RawApplication.ActiveCell);

        public IWorkbooksD Workbooks => DecoratorFactory.WorkbooksD(RawApplication.Workbooks);
        public string StatusBar { get => RawApplication.StatusBar?.ToString() ?? ""; set => RawApplication.StatusBar = value; }

        //Non-mvp: decorate window. It's only worth doing if we need to use it more than rarely.
        public Window ActiveWindow => RawApplication.ActiveWindow;

        //Non-mvp: do we wrap this in a forwarder/transposer in case it's on a transposed sheet?
        public IRangeD Selection => GetSelectionAsRange();

        public IWorksheetD ActiveSheetD => DecoratorFactory.WorksheetD((Worksheet) RawApplication.ActiveSheet);

        public bool ScreenUpdating
        {
            get => RawApplication.ScreenUpdating;
            set => RawApplication.ScreenUpdating = value;
        }

        public bool DisplayAlerts
        {
            get => RawApplication.DisplayAlerts;
            set => RawApplication.DisplayAlerts = value;
        }

        //Non-mvp: Refactor this out to a separate worker object (it clutters the application object a bit)
        public void RunWithScreenUpdateOff(System.Action action)
        {
            bool screenUpdatingOrig = ScreenUpdating;
            if (screenUpdatingOrig)
            {
                Log.Debug("Turning screen updating off before running action");
                ScreenUpdating = false;
            }
            else
            {
                Log.Debug("Application screen updating is already off. No need to switch it off.");
            }
            try
            {
                action.Invoke();
            }
            finally
            {
                if (screenUpdatingOrig)
                {
                    ScreenUpdating = screenUpdatingOrig;
                    Log.Debug("Turning screen updating back on after running action");
                }
            }
        }

        public void RunWithDisplayAlertsOff(System.Action actionToRun)
        {
            bool initialState = DisplayAlerts;
            try
            {
                DisplayAlerts = false;
                actionToRun();
            }
            finally
            {
                DisplayAlerts = initialState;
            }
        }

        private IRangeD GetSelectionAsRange() => InteropTypeValidator.GetMapValidate<Range, IRangeD>
            (() => RawApplication.Selection, DecoratorFactory.RangeD);

        public void QuitWithoutPrompt()
        {
            DisplayAlerts = false;
            RawApplication.Quit();
        }

        public void UnbreakableCalculate()
        {
            IApplicationTempState tempState = NewTempState();
            tempState.TempObject.CalculationInterruptKey = XlCalculationInterruptKey.xlNoKey;
            tempState.RunWithTempOptions(RawApplication.Calculate);
        }

        public void RunWithStatusBarMessage(System.Action action, string message)
        {
            IApplicationTempState tempState = NewTempState();
            tempState.TempObject.StatusBar = message;
            tempState.RunWithTempOptions(action);
        }
        public bool TryCompileVbaProject()
        {
            try
            {
                var vbe = RawApplication.VBE;
                vbe.MainWindow.Visible = true;
                dynamic controls = RawApplication.VBE.CommandBars["Menu Bar"].Controls["Debug"];
                controls.Controls["Compile VBAProject"].Execute();
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to compile VBA project", ex);
                return false;
            }
        }

        public IRangeD Range(string address)
        {
            try
            {
                Range rawRange = RawApplication.get_Range(address);                
                return DecoratorFactory.RangeD(rawRange);
            }
            catch(Exception e)
            {
                throw new InvalidOperationException("Interop Application object could not get range for reference formula " + address, e);
            }
        }

        public IRangeD Intersect(IRangeD range1, IRangeD range2)
        {
            Range rawRange1 = range1.RawRange;
            Range rawRange2 = range2.RawRange;
            Range rawIntersect = RawApplication.Intersect(rawRange1, rawRange2);
            return DecoratorFactory.RangeD(rawIntersect);
        }

        public IRangeD Union(IRangeD range1, IRangeD range2)
        {
            Range rawRange1 = range1.RawRange;
            Range rawRange2 = range2.RawRange;
            Range rawUnion = RawApplication.Union(rawRange1, rawRange2);
            return DecoratorFactory.RangeD(rawUnion);
        }

        public void ClearUndoList()
        {
            RawApplication.OnUndo("", "");
        }

        public void ClearCutCopyMode()
        {
            RawApplication.CutCopyMode = 0;
        }

        public IDictionary<string, IComAddinState> ComAddinStates
        {
            get => GetComAddinStates();
            set => SetComAddinStates(value);
        }

        private void SetComAddinStates(IDictionary<string, IComAddinState> states)
        {
            COMAddIns comAddIns = RawApplication.COMAddIns;
            foreach (KeyValuePair<string, IComAddinState> pair in states)
            {
                string addInId = pair.Key;
                COMAddIn? comAddin = GetCommAddInOrNull(comAddIns, addInId);
                if (comAddin != null)
                {
                    IComAddinState comAddinState = pair.Value;
                    SetComAddinState(comAddin, comAddinState);
                }
                else
                {
                    Log.Warn($"Error setting COM Addin state for addin with id: '{addInId}'");
                }
            }
        }

        private IDictionary<string, IComAddinState> GetComAddinStates()
        {
            COMAddIns comAddIns = RawApplication.COMAddIns;
            IDictionary<string, IComAddinState> addInStates
                = new Dictionary<string, IComAddinState>();
            foreach (object comAddInRaw in comAddIns)
            {
                if (comAddInRaw is COMAddIn comAddIn)
                {
                    string addinId = comAddIn.ProgId;
                    addInStates.Add(addinId, new ComAddinStatePoco()
                    {
                        AddinId = addinId,
                        IsConnected = comAddIn.Connect
                    });

                }
            }
            return addInStates;
        }

        private COMAddIn? GetCommAddInOrNull(COMAddIns comAddIns, string addInId)
        {
            try
            {
                return comAddIns.Item(addInId);
            }
            catch (Exception)
            {
                Log.Warn($"No COM AddIn found with name: {addInId}");
                return null;
            }
        }

        //Non-MVP: COMAddIn decorator for this
        private void SetComAddinState(COMAddIn comAddin, IComAddinState comAddinState)
        {
            try
            {
                comAddin.Connect = comAddinState.IsConnected;
            }
            catch (COMException e)
            {
                //COMException 0x80004005 can happen when trying to set connect property on "globally installed" AddIns- we can just ignore it
                Log.Debug("Error trying to set COM AddIn state for COM AddIn: " + comAddin.ProgId, e);
            }
        }
    }
}
