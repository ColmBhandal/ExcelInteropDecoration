using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using ExcelInteropDecoration.Decorator.application;

using ExcelInteropDecoration.Decorator.sheets;
using ExcelInteropDecoration.Decorator.vbComponent;
using ExcelInteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using ExcelInteropDecoration.Decorator._base;
using Microsoft.Office.Core;
using ExcelInteropDecoration.Decorator.range;
using CsharpExtras.Extensions;
using ExcelInteropDecoration.Decorator.names;
using CsharpExtras.Visitor;
using ExcelInteropDecoration.Decorator.util;
using CsharpExtras.Api;
using System.Runtime.InteropServices;
using CsharpExtras.IO;

namespace ExcelInteropDecoration.Decorator.workbook
{
    class WorkbookDImpl : DecoratorBase, IWorkbookD
    {
        public Workbook Workbook { get; }

        public WorkbookDImpl(IInteropDAPI api, Workbook workbook)
            : base(api)
        {
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        public string Path => Workbook.Path;

        public string Name => Workbook.Name;

        public string NameNoExtension => PathDecorator.GetFileNameWithoutExtension(Name);

        public string Extension => PathDecorator.GetExtension(Name);

        public ISheetsD Worksheets =>
            InteropTypeValidator.MapValidate<Sheets, ISheetsD>
            (Workbook.Worksheets, DecoratorFactory.WorksheetsD);

        public IEnumerable<IWorksheetD> WorksheetEnumerable() => Worksheets.WorksheetEnumerable();

        public IEnumerable<IWorksheetD> VisibleWorksheetsEnumerable() =>
            WorksheetEnumerable().Where(s => s.IsVisible);

        public IApplicationD Application => DecoratorFactory.ApplicationD(Workbook.Application);

        public IVBComponentD GetVbComponentByName(string vbCompName)
        {
            try
            {
                VBComponent vbComp = Workbook.VBProject.VBComponents.Item(vbCompName);
                return DecoratorFactory.VBComponentD(vbComp);
            }
            catch (IndexOutOfRangeException e)
            {
                throw new ArgumentException("VB Component not found: " + vbCompName, e);
            }
        }

        public void ImportVbModule(string filePath)
        {
            Workbook.VBProject.VBComponents.Import(filePath);
        }

        public void DeleteVbModule(IVBComponentD vbComponent)
        {
            Workbook.VBProject.VBComponents.Remove(vbComponent.RawVbComponent);
        }
        
        public void Save()
        {
            Workbook.Save();
        }

        public void Close()
        {
            Workbook.Close();
        }

        public void CloseWithoutPrompt()
        {
            Workbook.Close(false);
        }

        public void SaveAndClose()
        {
            Save(); CloseWithoutPrompt();
        }

        public void SaveAs(string fullName, XlFileFormat format = XlFileFormat.xlOpenXMLWorkbook)
        {
            try
            {
                Workbook.SaveAs(fullName, format);
            }
            catch(Exception e)
            {
                throw new InvalidOperationException
                    ($"Error while trying to save workbook {fullName} in file format {format.ToString()}", e);
            }
        }

        public void SaveAsHere(string newName)
        {
            XlFileFormat format = Workbook.FileFormat;
            SaveAs(System.IO.Path.Combine(Path, newName), format);
        }

        public IWorksheetD ActiveSheet() => DecoratorFactory.WorksheetD((Worksheet) Workbook.ActiveSheet);

        //non-mvp: we could write logic to record when the WB is closed and disallow access to the Workbook property by throwing an exception if so.
        public void SaveAndCloseWithoutRecalc()
        {
            //TODO: Refactor this to be like RunWithTempOptions, using an Application state object
            Microsoft.Office.Interop.Excel.Application app = Application.RawApplication;
            bool origCalcBeforeSave = app.CalculateBeforeSave;
            Application.RawApplication.CalculateBeforeSave = false;
            SaveAndClose();
            app.CalculateBeforeSave = origCalcBeforeSave;
        }

        public void ForceRecalculate()
        {
            Application.RawApplication.Calculate();
        }

        public void RecalculateIfNecessary()
        {
            if (IsRecalcPending())
            {
                ForceRecalculate();
            }
        }

        public bool IsRecalcPending()
        {
            return Application.RawApplication.CalculationState == XlCalculationState.xlPending;
        }

        public INamesD Names => DecoratorFactory.NamesD(Workbook.Names);

        public IRangeD Range(string sheetCellReference)
        {
            if(sheetCellReference.Length < 2)
            {
                throw new ArgumentException("Sheet-Cell reference cannot be less than 2 characters. Found: " + sheetCellReference);
            }

            (string sheetName, string address) = StringProcessor.SplitAddress(sheetCellReference);
            string fullAddress = StringProcessor.CombineAddress(Name, sheetName, address);

            return Application.Range(fullAddress);
        }

        public ISet<string> NamesAsStrings()
        {
            return NamesAsStringsAux(name => true);
        }

        public ISet<string> NamesWithErrorsAsStrings()
        {
            return NamesAsStringsAux(NameHasError);
        }

        private ISet<string> NamesAsStringsAux(Func<Name, bool> inclusionCondition)
        {
            ISet<string> nameSet = new HashSet<string>();
            try
            {
                //TODO: Use decorators for this
                Names names = Workbook.Names;
                foreach (Name name in names)
                {
                    if (inclusionCondition(name))
                    {
                        nameSet.Add(name.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn(string.Format("Exception while checking workbook '{0}' for name errors", NameNoExtension), ex);
            }
            return nameSet;
        }

        private bool NameHasError(Name name)
        {
            string refersTo = (string)name.RefersTo;
            return refersTo.Contains("#REF!") || name.Value == "#REF!";
        }

        public IWorksheetD? GetWorksheetDByNameOrNull(string sheetName) =>
            Worksheets.GetWorksheetDByNameOrNull(sheetName);


        public IRangeD NamedRange(string rangeName)
        {
            Names names = Workbook.Names;
            Name name;
            try
            {
                name = names.Item(rangeName);
            }
            catch (Exception e)
            {
                throw new InvalidOperationException(string.Format(
                    "Problem getting name {0} from workbook {1}", rangeName, NameNoExtension), e);
            }
            object refersTo = name.RefersTo;
            if (refersTo is string referenceFormula)
            {
                if (referenceFormula.StartsWith("="))
                {
                    string sheetCellReference = referenceFormula.Substring(1);
                    return Range(sheetCellReference);
                }
                throw new InvalidOperationException(string.Format(
                    "Reference formula {0} does not start with = as expected. Named Range {1} cannot be retrieved.",
                    referenceFormula, rangeName));
            }
            throw new InvalidCastException(string.Format(
                "Interop returned a non-string type for range name {0} in workbook {1}. The type returned was {2}",
                rangeName, NameNoExtension, refersTo.GetType()));
        }

        public IRangeD? NamedRangeOrNull(string rangeName)
        {
            try
            {
                return NamedRange(rangeName);
            }
            catch (Exception)
            {
                Log.Debug($"Range name {rangeName} not found in workbook {NameNoExtension}. Returning null.");
                return null;
            }
        }

    }
}
