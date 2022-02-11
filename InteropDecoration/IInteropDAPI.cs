using InteropDecoration.Decorator.application;
using InteropDecoration._base;
using InteropDecoration.Decorator;
using InteropDecoration.Helper.TempState;
using InteropDecoration.Helper.Validation;
using InteropDecoration.Decorator.range;
using InteropDecoration.Helper.ColourDataProcessor;
using InteropDecoration.Decorator.util;
using InteropDecoration.Helper.TempState.Application;

namespace InteropDecoration
{
    public interface IInteropDAPI
    {
        /// <summary>
        /// Set this property to define the type of loggers used by this instance of the API.
        /// </summary>
        Func<Type, ILogger> LoggerGenerator { get; set; }

        IDecoratorFactory NewDecoratorFactory();
        IApplicationTempState NewApplicationTempState(IApplicationD application);
        ILogger NewLogger(Type type);
        IInteropTypeValidator NewInteropTypeValidator();
        IRangeDataTransformer NewRangeDataTransformer();
        IColourDataProcessor NewColourDataProcessor();
        IInteropStringProcessor NewInteropStringProcessor();
    }
}