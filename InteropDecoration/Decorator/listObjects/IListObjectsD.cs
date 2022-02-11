using CsharpExtras.ValidatedType.Numeric.Integer;
using InteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;

namespace InteropDecoration.Decorator.listObjects
{
    public interface IListObjectsD
    {
        IListObjectD this[string tableName] { get; }

        ListObjects RawListObjects { get; }

        bool HasTable(string tableName);
    }
}