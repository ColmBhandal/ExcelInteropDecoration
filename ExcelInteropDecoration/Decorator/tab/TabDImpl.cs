using ExcelInteropDecoration.Decorator._base;

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Decorator.tab
{
    class TabDImpl : DecoratorBase, ITabD
    {
        private readonly Tab _rawTab;

        public TabDImpl(IInteropDAPI api, Tab rawTab) : base(api)
        {
            _rawTab = rawTab ?? throw new ArgumentNullException(nameof(rawTab));
        }
        public int? ColourRgb
        {
            get => GetTabColourOrNull();
            set => SetTabColour(value);
        }

        private void SetTabColour(int? colourRgbOrNull)
        {
            if(colourRgbOrNull == null)
            {
                //Interop appears to use the boolean value false to indicate the tab colour isn't set
                _rawTab.Color = false;
            }
            else
            {
                int colourRgb = colourRgbOrNull.Value;
                int colourBgr = ColourDataProcessor.RgbColourToBgr(colourRgb);
                _rawTab.Color = colourBgr;
            }
        }

        private int? GetTabColourOrNull()
        {
            try
            {
                int bgrColour = (int)_rawTab.Color;
                int rgbColour = ColourDataProcessor.BgrColourToRgb(bgrColour);
                return rgbColour;
            }
            catch(InvalidCastException)
            {
                Log.Debug("Tab colour is not an integer. This means the tab colour is not set. Returning null from the decoration layer.");
                return null;
            }
        }

        public void FillRgb(int rgbColour)
        {
            ColourRgb = rgbColour;
        }
    }
}
