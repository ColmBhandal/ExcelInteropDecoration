namespace ExcelInteropDecoration.Helper.ColourDataProcessor
{
    public interface IColourDataProcessor
    {
        int[,] RgbColourToBgr(int[,] rgbArr);
        int[,] BgrColourToRgb(int[,] bgrArr);

        int RgbColourToBgr(int rgb);
        int BgrColourToRgb(int bgr);
        string RgbColourToHex(int rgb);
        int HexToRgbColour(string hex);
    }
}