using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelInteropDecoration.Decorator.range
{
    public interface IRangeDataTransformer
    {
        /// <summary>
        /// Will be invoked before a change is made to the given range
        /// </summary>
        event Action<Range> BeforeChange;
        int[,] ReadBackgroundColours(Microsoft.Office.Interop.Excel.Range range);
        string[,] ReadFormulas(Microsoft.Office.Interop.Excel.Range range);
        string[,] ReadValues(Microsoft.Office.Interop.Excel.Range range);
        void WriteBackgroundColours(Microsoft.Office.Interop.Excel.Range range, int[,] colours);
        void WriteFormulas(Microsoft.Office.Interop.Excel.Range range, string[,] formulas);
        void WriteValues(Microsoft.Office.Interop.Excel.Range range, string[,] values);
    }
}