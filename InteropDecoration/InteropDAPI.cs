using InteropDecoration.Decorator.application;
using InteropDecoration.Decorator.workbook;
using InteropDecoration._base;
using InteropDecoration.Decorator;
using InteropDecoration.Helper.TempState;
using Microsoft.Office.Interop.Excel;
using InteropDecoration.Helper.Validation;
using InteropDecoration.Decorator.range;
using InteropDecoration.Helper.ColourDataProcessor;
using InteropDecoration.Helper.TempState.Application;
using InteropDecoration.Decorator.util;

namespace InteropDecoration
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