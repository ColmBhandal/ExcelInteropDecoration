using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Helper.ColourDataProcessor
{
    internal class ColourDataProcessorImpl : IColourDataProcessor
    {
        public int BgrColourToRgb(int bgr)
        {
            int r = bgr % 256;
            int g = (bgr / 256) % 256;
            int b = bgr / (256 * 256);

            return r * (256 * 256) + g * 256 + b;
        }

        public int[,] BgrColourToRgb(int[,] bgrArr)
        {
            for (int i = 0; i < bgrArr.GetLength(0); i++)
            {
                for (int j = 0; j < bgrArr.GetLength(1); j++)
                {
                    bgrArr[i, j] = BgrColourToRgb(bgrArr[i, j]);
                }
            }
            return bgrArr;
        }

        public int HexToRgbColour(string hex)
        {
            return int.Parse(hex, NumberStyles.AllowHexSpecifier);
        }

        public int RgbColourToBgr(int rgb)
        {
            int b = rgb % 256;
            int g = (rgb / 256) % 256;
            int r = rgb / (256 * 256);

            return b * (256 * 256) + g * 256 + r;
        }

        public int[,] RgbColourToBgr(int[,] rgbArr)
        {
            for (int i = 0; i < rgbArr.GetLength(0); i++)
            {
                for (int j = 0; j < rgbArr.GetLength(1); j++)
                {
                    rgbArr[i, j] = RgbColourToBgr(rgbArr[i, j]);
                }
            }
            return rgbArr;
        }

        public string RgbColourToHex(int rgb)
        {
            return rgb.ToString("X");
        }
    }
}
