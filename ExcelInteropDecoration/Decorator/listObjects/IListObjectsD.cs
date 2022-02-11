using CsharpExtras.ValidatedType.Numeric.Integer;
using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropDecoration.Decorator.listObjects
{
    public interface IListObjectsD
    {
        IListObjectD this[string tableName] { get; }

        ListObjects RawListObjects { get; }

        bool HasTable(string tableName);
    }
}