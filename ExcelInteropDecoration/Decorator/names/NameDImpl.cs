using ExcelInteropDecoration.Decorator._base;
using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelInteropDecoration.Decorator.names
{
    class NameDImpl : DecoratorBase, INameD
    {
        public Name RawName { get; }

        //Excel seems to allow some Name objects that don't have any refers to range
        //It seems that a Name is more general than a named range
        //That is, some, but not all Names correspond to named ranges
        //Thus, only some names will actually have a refers to range
        //In the case where a name doesn't have a refers to range, the property "fails" https://docs.microsoft.com/en-us/office/vba/api/excel.name.referstorange
        //In that case we just catch the exception here and return null
        //Null indicates that this Name is not a named range
        public IRangeD? RefersToRangeOrNull => InteropTypeValidator.GetMapValidateOrNull<Range, IRangeD>
            (() => RawName.RefersToRange, DecoratorFactory.RangeD);

        public string Name => RawName.Name;

        public NameDImpl(IInteropDAPI api, Name rawName) : base(api)
        {
            RawName = rawName ?? throw new ArgumentNullException(nameof(rawName));
        }

        public void Delete()
        {
            try
            {
                RawName.Delete();
            }
            catch(Exception e)
            {
                throw new InvalidOperationException("Exception thrown trying to delete range name", e);
            }
        }
    }
}
