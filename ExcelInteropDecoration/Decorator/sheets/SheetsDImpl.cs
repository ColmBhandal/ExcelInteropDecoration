using ExcelInteropDecoration.Decorator._base;
using ExcelInteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelInteropDecoration.Decorator.sheets
{
    class SheetsDImpl : DecoratorBase, ISheetsD
    {
        public SheetsDImpl(IInteropDAPI api, Sheets worksheets) : base(api)
        {
            Worksheets = worksheets;
        }
        public int Count => Worksheets.Count;

        public Sheets Worksheets { get; private set; }

        public IEnumerable<IWorksheetD> WorksheetEnumerable()
        {
            foreach(object? rawSheet in Worksheets)
            {
                //Note: the sheets object is mix of charts and worksheets
                //See: https://docs.microsoft.com/en-us/office/vba/api/excel.sheets
                //Therefore, we can't force-cast or fail if we find a non-Worksheet object- it could be a chart
                //Nor should we log a debug message saying "non-Worksheet found" or something - it'll just add noise if there are charts present
                if (rawSheet is Worksheet sheet)
                {
                    yield return DecoratorFactory.WorksheetD(sheet);
                }                
            }
        }

        public IWorksheetD this[string index]
        {
            get
            {
                object rawWorksheetObject;
                try
                {
                    rawWorksheetObject = Worksheets[index];
                }
                catch (Exception ex)
                {
                     throw new ArgumentException(string.Format("Could not get sheet by string value index: '{0}'", index), ex);
                }
                return InteropTypeValidator.MapValidate<Worksheet, IWorksheetD>(rawWorksheetObject, DecoratorFactory.WorksheetD);                
            }
        }

        public IWorksheetD AddNewSheet(string sheetName)
        {
            Worksheet worksheet = (Worksheet)Worksheets.Add();
            worksheet.Name = sheetName;
            return DecoratorFactory.WorksheetD(worksheet);
        }

        public bool HasSheet(string sheetName)
        {
            //non-mvp: is there a better way of checking for sheet existence than attempting to index and catching an exception in the false case?
            try
            {
                _ = Worksheets[sheetName];
                return true;
            }
            catch(Exception)
            {
                Log.Debug($"Expected exception caught trying to index sheet name {sheetName} - implies sheet does not exist.");
                return false;
            }
        }

        public IWorksheetD? GetWorksheetDByNameOrNull(string sheetName)
        {
            if (HasSheet(sheetName))
            {
                return this[sheetName];
            }
            return null;
        }
    }
}
