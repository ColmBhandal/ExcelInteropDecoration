using InteropDecoration._exception;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Helper.Validation
{
    //TODO: Add unit tests: happy, failing generator returns null, failing mapper returns null
    internal class InteropTypeValidatorImpl : IInteropTypeValidator
    {
        public TReturn? GetMapValidateOrNull<TInterop, TReturn>
            (Func<object?> sourceGenerator, Func<TInterop, TReturn> mapper)
            where TReturn : class
        {
            object? sourceObject;
            try
            {
                sourceObject = sourceGenerator();
            }
            catch (Exception)
            {
                return null;
            }
            try
            {
                return MapValidate(sourceObject, mapper);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public TReturn GetMapValidate<TInterop, TReturn>(Func<object?> sourceGenerator, Func<TInterop, TReturn> mapper)
        {
            object? sourceObject;
            try
            {
                sourceObject = sourceGenerator();
            }
            catch (Exception ex)
            {
                Type expectedType = typeof(TInterop);
                throw new InteropObjectGenerationException(
                    $"Could not generate source object of type {expectedType}. Source object generation failed.", ex);
            }
            return MapValidate(sourceObject, mapper);
        }

        public TReturn MapValidate<TInterop, TReturn>(object? sourceObject, Func<TInterop, TReturn> mapper)
        {
            if (sourceObject is TInterop interopObject)
            {
                return mapper(interopObject);
            }
            Type expectedType = typeof(TInterop);
            if (sourceObject == null)
            {
                throw new InteropTypeCastException($"Expected a type of {expectedType.FullName} but found a null object.");
            }
            throw new InteropTypeCastException(expectedType, sourceObject.GetType());
        }
    }
}
