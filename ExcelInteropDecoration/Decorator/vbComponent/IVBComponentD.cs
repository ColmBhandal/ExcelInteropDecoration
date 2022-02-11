using ExcelInteropDecoration._base;
using ExcelInteropDecoration.Decorator.workbook;
using Microsoft.Vbe.Interop;

namespace ExcelInteropDecoration.Decorator.vbComponent
{
    public interface IVBComponentD
    {
        VBComponent RawVbComponent { get; }

        VBComponentType GetComponentType();

        string GetComponentRawName();

        string? GetComponentPrettyNameOrNull();

        string GetVbCodeLines(int numberOfLines);

        void DeleteVbCodeLines(int numberOfLines);

        void DeleteAllCode();

        int CountCodeLines();

        void ImportCodeFromFile(string filePath);

        void ExportCodeToFile(string filePath);
    }

    public interface IVBComponentDBuilder : IBuilder<IVBComponentD>
    {
        IVBComponentDBuilder WithWorkbook(IWorkbookD workbook);

        IVBComponentDBuilder WithComponentName(string vbCompName);

        IVBComponentDBuilder WithVbComponent(VBComponent vbComp);
    }

    public interface IVBComponentDBuilderFactory : IFactory<IVBComponentDBuilder>
    {
    }
}
