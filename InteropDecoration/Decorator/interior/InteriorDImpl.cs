using InteropDecoration.Decorator._base;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Decorator.interior
{
    internal class InteriorDImpl : DecoratorBase, IInteriorD
    {
        Interior RawInterior { get; }

        public InteriorDImpl(IInteropDAPI interopDApi, Interior rawInterior) : base(interopDApi)
        {
            RawInterior = rawInterior ?? throw new ArgumentNullException(nameof(rawInterior));
        }

        public int ColourBgr
        {
            //TODO: Add safer casting
            get => (int)(double)RawInterior.Color;
            set => RawInterior.Color = value;
        }
    }
}
