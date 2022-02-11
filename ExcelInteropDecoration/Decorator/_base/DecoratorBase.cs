using CsharpExtras.IO;
using ExcelInteropDecoration._base;
using ExcelInteropDecoration.Decorator.util;
using ExcelInteropDecoration.Helper.ColourDataProcessor;
using ExcelInteropDecoration.Helper.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Decorator._base
{
    internal class DecoratorBase : BaseClass
    {
        public DecoratorBase(IInteropDAPI interopDApi) : base(interopDApi)
        {
        }

        private IDecoratorFactory? _decoratorFactory;
        protected IDecoratorFactory DecoratorFactory => _decoratorFactory??= InteropDApi.NewDecoratorFactory();
        private IInteropTypeValidator? _interopTypeValidator;
        protected IInteropTypeValidator InteropTypeValidator =>
            _interopTypeValidator ??= InteropDApi.NewInteropTypeValidator();

        private IPathDecorator? _pathDecorator;
        protected IPathDecorator PathDecorator =>
            _pathDecorator ??= CsharpExtrasApi.NewPathDecorator();

        private IInteropStringProcessor? _stringProcessor;
        protected IInteropStringProcessor StringProcessor => 
            _stringProcessor ??= InteropDApi.NewInteropStringProcessor();

        private IColourDataProcessor? _colourDataProcessor;
        protected IColourDataProcessor ColourDataProcessor => _colourDataProcessor
            ??= InteropDApi.NewColourDataProcessor();
    }
}
