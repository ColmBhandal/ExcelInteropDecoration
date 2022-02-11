using ExcelInteropDecoration.Decorator.application;
using ExcelInteropDecoration.Decorator.workbook;
using ExcelInteropDecoration._base;
using ExcelInteropDecoration.Decorator;
using ExcelInteropDecoration.Helper.TempState;
using Microsoft.Office.Interop.Excel;
using ExcelInteropDecoration.Helper.Validation;
using ExcelInteropDecoration.Decorator.range;
using ExcelInteropDecoration.Helper.ColourDataProcessor;
using ExcelInteropDecoration.Helper.TempState.Application;
using ExcelInteropDecoration.Decorator.util;

namespace ExcelInteropDecoration
{
    public class InteropDAPI : IInteropDAPI
    {
        //TODO: Test that 2 InteropDAPIs can co-exist with different logger generators i.e. the test will ensure this
        public Func<Type, ILogger> LoggerGenerator { get; set; } = 
            t => new ConsoleLoggerImpl();

        public IDecoratorFactory NewDecoratorFactory() => new DecoratorFactoryImpl(this);

        public ILogger NewLogger(Type type) => LoggerGenerator(type);

        public IApplicationTempState NewApplicationTempState(IApplicationD application) =>
            new ApplicationTempStateImpl(application);

        public IInteropTypeValidator NewInteropTypeValidator() =>
            new InteropTypeValidatorImpl();

        public IRangeDataTransformer NewRangeDataTransformer() =>
            new RangeDataTransformerImpl(this);

        public IColourDataProcessor NewColourDataProcessor() =>
            new ColourDataProcessorImpl();

        public IInteropStringProcessor NewInteropStringProcessor() =>
            new InteropStringProcessorImpl(this);
    }
}