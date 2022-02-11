
namespace ExcelInteropDecoration.Helper.Validation
{
    public interface IInteropTypeValidator
    {
        /// <summary>
        /// Applies the given mapper to the source object if it is of the correct type, otherwise throws an exception
        /// </summary>
        /// <typeparam name="TInterop">Expected type of the source object</typeparam>
        /// <typeparam name="TReturn">Type to return</typeparam>
        /// <param name="sourceObject">Source object to map</param>
        /// <param name="mapper">Function to apply to source object to produce return value</param>
        /// <returns></returns>
        TReturn MapValidate<TInterop, TReturn>(object? sourceObject, Func<TInterop, TReturn> mapper);

        /// <summary>
        /// Generates a source object then applies the given mapper to the source object if it is of the correct type, otherwise throws an exception
        /// </summary>
        /// <typeparam name="TInterop">Expected type of the source object</typeparam>
        /// <typeparam name="TReturn">Type to return</typeparam>
        /// <param name="sourceGenerator">Function to generate the source object to be mapped.</param>
        /// <param name="mapper">Function to apply to source object to produce return value</param>        
        /// <returns></returns>
        TReturn GetMapValidate<TInterop, TReturn>(Func<object?> sourceGenerator, Func<TInterop, TReturn> mapper);

        /// <summary>
        /// Generates a source object then applies the given mapper. If any exceptions are thrown, returns null.
        /// </summary>
        /// <typeparam name="TInterop">Expected type of the source object</typeparam>
        /// <typeparam name="TReturn">Type to return</typeparam>
        /// <param name="sourceGenerator">Function to generate the source object to be mapped.</param>
        /// <param name="mapper">Function to apply to source object to produce return value</param>        
        /// <returns></returns>
        TReturn? GetMapValidateOrNull<TInterop, TReturn>(Func<object?> sourceGenerator, Func<TInterop, TReturn> mapper) where TReturn : class;
    }
}