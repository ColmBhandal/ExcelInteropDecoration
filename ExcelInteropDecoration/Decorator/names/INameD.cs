using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropDecoration.Decorator.names
{
    public interface INameD
    {
        Name RawName { get; }
        IRangeD? RefersToRangeOrNull { get; }
        string Name { get; }

        void Delete();
    }
}