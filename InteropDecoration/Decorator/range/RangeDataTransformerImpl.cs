using InteropDecoration._base;
using PerformanceRecorder.Attribute;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace InteropDecoration.Decorator.range
{
    [PerformanceLogging]
    class RangeDataTransformerImpl : BaseClass, IRangeDataTransformer
    {
        public RangeDataTransformerImpl(IInteropDAPI interopDApi) : base(interopDApi)
        {
        }

        //Non-MVP: Convert raw range type to a wrapped type in this event
        public event Action<Range>? BeforeChange;

        public string[,] ReadFormulas(Range range)
        {
            return ConvertToStringArray2D(GetArray(range, () => range.Formula));
        }

        public string[,] ReadValues(Range range)
        {
            return ConvertToStringArray2D(GetArray(range, () => range.Value2));
        }

        public int[,] ReadBackgroundColours(Range range)
        {
            if (range.Areas.Count > 1)
            {
                throw new ArgumentException(string.Format(
                    "Cannot read background colours for non-rectangular range {0}", range.AddressLocal));
            }
            int[,] colours = new int[range.Rows.Count, range.Columns.Count];
            for (int r = 1; r <= range.Rows.Count; r++)
            {
                for (int c = 1; c <= range.Columns.Count; c++)
                {
                    colours[r - 1, c - 1] = (int)(double)((Range)range.Cells[r, c]).Interior.Color;
                }
            }
            return colours;
        }

        public void WriteFormulas(Range range, string[,] formulas)
        {
            PrepareChangeHandlerForChange(range);
            range.Formula = FormulaStringArrayToObjectArray(formulas);
        }

        public void WriteValues(Range range, string[,] values)
        {
            PrepareChangeHandlerForChange(range);
            range.Value2 = FormulaStringArrayToObjectArray(values);
        }

        public void WriteBackgroundColours(Range range, int[,] colours)
        {
            for (int r = 1; r <= range.Rows.Count; r++)
            {
                for (int c = 1; c <= range.Columns.Count; c++)
                {
                    ((Range)range.Cells[r, c]).Interior.Color = colours[r - 1, c - 1];
                }
            }
        }

        //Non-MVP: Wrap the range type
        private void PrepareChangeHandlerForChange(Range range)
        {
            BeforeChange?.Invoke(range);
        }

        private Array GetArray(Range range, Func<object> dataProducer)
        {
            if (range.Cells.Count == 0)
            {
                return Array.CreateInstance(typeof(Object), 0, 0);
            }
            if (range.Cells.Count == 1)
            {
                //Creating a 1-based array to fit with Excel's 1-based indexing
                Array retArray = OneBasedSingletonArray2D();
                retArray.SetValue(dataProducer.Invoke(), 1, 1);
                return retArray;
            }
            else
            {
                return (Array)dataProducer.Invoke();
            }
        }

        //TODO: Write a test for this
        //TODO: Verify you don't get an index out of bounds exception when you index values at the max row, col
        //non-mvp: move to a util somewhere
        ///NB: The parameter values is 1-based but the return type is 0-based
        private string[,] ConvertToStringArray2D(Array values)
        {
            string[,] retArray = new string[values.GetLength(0), values.GetLength(1)];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    string? strValueOrNull = values.GetValue(row, col)?.ToString();                    
                    if(strValueOrNull == null)
                    {
                        Log.Debug($"Null value found at (row, col) = ({row},{col}). Converting to an empty string.");
                    }
                    retArray[row - 1, col - 1] = strValueOrNull ?? "";
                }
            }

            return retArray;
        }

        private Array OneBasedSingletonArray2D()
        {
            return Array.CreateInstance(typeof(Object), new int[] { 1, 1 }, new int[] { 1, 1 });
        }

        private object[,] FormulaStringArrayToObjectArray(string[,] strFormulas)
        {
            object[,] objFormulas = new object[strFormulas.GetLength(0), strFormulas.GetLength(1)];
            for (int i = 0; i < strFormulas.GetLength(0); i++)
            {
                for (int j = 0; j < strFormulas.GetLength(1); j++)
                {
                    objFormulas[i, j] = strFormulas[i, j];
                }
            }
            return objFormulas;
        }
    }
}
