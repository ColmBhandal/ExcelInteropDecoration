using InteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;

namespace InteropDecoration.Decorator.names
{
    public interface INameD
    {
        Name RawName { get; }
        IRangeD? RefersToRangeOrNull { get; }
        string Name { get; }

        void Delete();
    }
}