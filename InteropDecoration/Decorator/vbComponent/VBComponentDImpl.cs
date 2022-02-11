using InteropDecoration.Decorator._base;
using Microsoft.Vbe.Interop;

namespace InteropDecoration.Decorator.vbComponent
{
    class VBComponentDImpl : DecoratorBase, IVBComponentD
    {
        public VBComponent RawVbComponent { get; }
        private const int SheetNamePropertyIndex = 7;

        public VBComponentDImpl(IInteropDAPI api, VBComponent vbComponent) : base(api)
        {
            RawVbComponent = vbComponent;
        }

        public int CountCodeLines()
        {
            return RawVbComponent.CodeModule.CountOfLines;
        }

        public void DeleteAllCode()
        {
            RawVbComponent.CodeModule.DeleteLines(1, CountCodeLines());
        }

        public void DeleteVbCodeLines(int numberOfLines)
        {
            RawVbComponent.CodeModule.DeleteLines(1, numberOfLines);
        }

        public void ExportCodeToFile(string filePath)
        {
            RawVbComponent.Export(filePath);
        }

        public string? GetComponentPrettyNameOrNull()
        {
            return RawVbComponent.Properties.Item(SheetNamePropertyIndex).Value.ToString();
        }

        public string GetComponentRawName()
        {
            return RawVbComponent.Name;
        }

        public VBComponentType GetComponentType()
        {
            string rawType = RawVbComponent.Type.ToString();
            switch (rawType)
            {
                case "vbext_ct_ClassModule": 
                    return VBComponentType.VBCompTypeClassModule;
                case "vbext_ct_StdModule":
                    return VBComponentType.VBCompTypeStdModule;
                case "vbext_ct_Document":
                    return VBComponentType.VBCompTypeDocument;
                case "vbext_ct_MSForm":
                    return VBComponentType.VBCompTypeForm;
                default:
                    throw new ArgumentException($"Could not find matching component type for raw type: {rawType}");
            }
        }

        public string GetVbCodeLines(int numberOfLines)
        {
            return RawVbComponent.CodeModule.Lines[1, numberOfLines];
        }

        public void ImportCodeFromFile(string filePath)
        {
            RawVbComponent.CodeModule.AddFromFile(filePath);
        }
        
    }
}