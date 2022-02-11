using ExcelInteropDecoration.Decorator.application;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Helper.TempState.Application
{
    internal class ApplicationTempStateImpl : TempStateBase<IApplicationState>, IApplicationTempState
    {
        private IApplicationD _applicationDecorator { get; }
        public ApplicationTempStateImpl(IApplicationD permanentObject)
            : base(new ApplicationStateWrapper(permanentObject))
        {
            _applicationDecorator = permanentObject ?? throw new ArgumentNullException(nameof(permanentObject));
        }

        protected override void Copy(IApplicationState from, IApplicationState to)
        {
            if (from.CalculateBeforeSaveOrNull != null)
            {
                to.CalculateBeforeSave = from.CalculateBeforeSaveOrNull.Value;
            }
            if (from.VisibleOrNull != null)
            {
                to.Visible = from.VisibleOrNull.Value;
            }                            
            if (from.StatusBar != null)
            {
                to.StatusBar = from.StatusBar;
            }
            to.CalculationInterruptKey = from.CalculationInterruptKey;
            to.ComAddinStates = from.ComAddinStates;
        }

        protected override IApplicationState NewPoco()
        {
            return new ApplicationStatePoco();
        }

        public void RunWithTempOptions(Action<IApplicationD> action)
        {
            RunWithTempOptions(() => action.Invoke(_applicationDecorator));
        }
    }
}
