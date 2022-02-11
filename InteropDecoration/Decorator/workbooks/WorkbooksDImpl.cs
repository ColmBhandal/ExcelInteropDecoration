using CsharpExtras.Enumerable.NonEmpty;
using InteropDecoration.Decorator._base;
using InteropDecoration.Decorator.workbook;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace InteropDecoration.Decorator.workbooks
{
    class WorkbooksDImpl : DecoratorBase, IWorkbooksD
    {
        public Workbooks RawWorkbooks { get; private set; }

        public WorkbooksDImpl(IInteropDAPI api, Workbooks rawWorkbooks) : base(api)
        {
            RawWorkbooks = rawWorkbooks;
        }

        //non-MVP: add a version of this function which just takes a single sheet name
        //and maybe add a version that just takes a collection and throws an exception if the collection is empty (though the current way enforces non-emptiness)
        public IWorkbookD AddWorkbookWithSheets(INonEmptyEnumerable<string> sheetNames)
        {
            Workbook workbook = AddWorkbookWithSheets(sheetNames.Count);
            //non-MVP: add method to do cast to Worksheet on the WB decorator. Same with cast in loop below.
            
            int index = 1;
            foreach (string sheetName in sheetNames)
            {
                Worksheet sheet = (Worksheet)workbook.Worksheets[index];
                sheet.Name = sheetName;
                index++;
            }
            return DecoratorFactory.WorkbookD(workbook);            
        }

        public IWorkbookD Open(string filePath)
        {
            string fileFullPath = Path.GetFullPath(filePath);
            bool doesFileExist = File.Exists(fileFullPath);
            if (!doesFileExist)
            {
                throw new FileNotFoundException($"File not found: fileFullPath");
            }
            try
            {
                Workbook workbook;
                workbook = RawWorkbooks.Open(filePath);
                return DecoratorFactory.WorkbookD(workbook);
            }
            catch(COMException e)
            {
                throw new InvalidOperationException($"COM Exception thrown when trying to open workbook", e);
            }
        }

        private Workbook AddWorkbookWithSheets(int sheetCount)
        {
            int origSheetsInNewWorkbook = RawWorkbooks.Application.SheetsInNewWorkbook;
            RawWorkbooks.Application.SheetsInNewWorkbook = sheetCount;
            Workbook workbook = RawWorkbooks.Add();
            RawWorkbooks.Application.SheetsInNewWorkbook = origSheetsInNewWorkbook;
            return workbook;
        }
    }
}
