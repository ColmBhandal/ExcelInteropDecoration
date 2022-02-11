using ExcelInteropDecoration.Decorator.application;
using ExcelInteropDecoration._base;
using ExcelInteropDecoration.Decorator;
using ExcelInteropDecoration.Helper.TempState;
using ExcelInteropDecoration.Helper.Validation;
using ExcelInteropDecoration.Decorator.range;
using ExcelInteropDecoration.Helper.ColourDataProcessor;
using ExcelInteropDecoration.Decorator.util;
using ExcelInteropDecoration.Helper.TempState.Application;

namespace ExcelInteropDecoration
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