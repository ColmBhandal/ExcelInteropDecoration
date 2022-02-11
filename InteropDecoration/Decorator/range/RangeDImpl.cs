using InteropDecoration.Decorator.util;
using InteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using CsharpExtras.Enumerable.OneBased;
using PerformanceRecorder.Attribute;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Schema;
using InteropDecoration.Decorator.names;
using InteropDecoration.Decorator.application;
using InteropDecoration.Decorator.comment;
using InteropDecoration.Decorator._base;
using Range = Microsoft.Office.Interop.Excel.Range;
using InteropDecoration.Helper.ColourDataProcessor;
using System.Runtime.InteropServices;
using InteropDecoration.Decorator.interior;
using CsharpExtras.Extensions;
using static CsharpExtras.Extensions.ArrayOrientationClass;

namespace InteropDecoration.Decorator.range
{
    class RangeDImpl : DecoratorBase, IRangeD
    {
        public RangeDImpl(IInteropDAPI api, Range range) : base(api)
        {
            RawRange = range;
        }
        private IRangeDataTransformer? _rangeDataTransformer;
        protected IRangeDataTransformer RangeDataTransformer => _rangeDataTransformer
            ??= InteropDApi.NewRangeDataTransformer();

        //TODO: Put these obsolete markers on all of the "Raw" Interop properties underlying the entire decoration layer
        [Obsolete("Don't use raw Excel Interop unless strictly necessary - use decoration layer instead",
            false)]
        public Range RawRange { get; }

        public IApplicationD Application => Parent.Application;
        public IRangeD EntireRow => DecoratorFactory.RangeD(RawRange.EntireRow);
        public IRangeD EntireColumn => DecoratorFactory.RangeD(RawRange.EntireColumn);

        public int CellCount => RawRange.Cells.Count;

        public bool IsEmpty => CellCount == 0;
        public bool IsSingleCell => CellCount == 1;

        public ICommentD? CommentOrNull
            => RawRange.Comment == null ? null : DecoratorFactory.CommentD(RawRange.Comment);

        public IOneBasedArray2D<int> RgbBackgroundColours
        {
            get => CsharpExtrasApi.NewOneBasedArray2D(RgbBackgroundColoursZeroBased);
            set => RgbBackgroundColoursZeroBased = value.ZeroBasedEquivalent;
        }

        public IOneBasedArray<int> FirstRowRgbColour
        {
            get => CsharpExtrasApi.NewOneBasedArray(FirstRowRgbColourZeroBased);
            set => FirstRowRgbColourZeroBased = value.ZeroBasedEquivalent;
        }

        public IOneBasedArray<int> FirstColumnRgbColour
        {
            get => CsharpExtrasApi.NewOneBasedArray(FirstColumnRgbColourZeroBased);
            set => FirstColumnRgbColourZeroBased = value.ZeroBasedEquivalent;
        }

        private IRangeD FirstColumn() => ColumnAt(1);

        private IRangeD FirstRow() => RowAt(1);

        public string[,] ValueZeroBased
        {
            get => RangeDataTransformer.ReadValues(RawRange);
            set => RangeDataTransformer.WriteValues(RawRange, value);
        }

        public IOneBasedArray2D<string> Value
        {
            get => CsharpExtrasApi.NewOneBasedArray2D(ValueZeroBased);
            set => ValueZeroBased = value.ZeroBasedEquivalent;
        }

        public string[,] FormulaZeroBased
        {
            get => RangeDataTransformer.ReadFormulas(RawRange);
            set => RangeDataTransformer.WriteFormulas(RawRange, value);
        }

        public IOneBasedArray2D<string> Formula
        {
            get => CsharpExtrasApi.NewOneBasedArray2D(FormulaZeroBased);
            set => FormulaZeroBased = value.ZeroBasedEquivalent;
        }

        public string FirstCellValue
        {
            get => FirstCell.ValueZeroBased[0, 0];
            set => FirstCell.ValueZeroBased = new string[,] { { value } };
        }

        public string FirstCellFormula
        {
            get => FirstCell.FormulaZeroBased[0, 0];
            set => FirstCell.FormulaZeroBased = new string[,] { { value } };
        }

        public string[] FirstRowValueZeroBased
        {
            get => FirstRow().ValueZeroBased.SliceRow(0);
            set => FirstRow().ValueZeroBased = value.To2DArray(ArrayOrientation.COLUMN);
        }

        public string[] FirstColumnValueZeroBased
        {
            get => FirstColumn().ValueZeroBased.SliceColumn(0);
            set => FirstColumn().ValueZeroBased = value.To2DArray(ArrayOrientation.ROW);
        }

        public string[] FirstRowFormulaZeroBased
        {
            get => FirstRow().FormulaZeroBased.SliceRow(0);
            set => FirstRow().FormulaZeroBased = value.To2DArray(ArrayOrientation.COLUMN);
        }

        public string[] FirstColumnFormulaZeroBased
        {
            get => FirstColumn().FormulaZeroBased.SliceColumn(0);
            set => FirstColumn().FormulaZeroBased = value.To2DArray(ArrayOrientation.ROW);
        }


        public int[,] RgbBackgroundColoursZeroBased
        {
            get => ColourDataProcessor.BgrColourToRgb(RangeDataTransformer.ReadBackgroundColours(RawRange));
            set => RangeDataTransformer.WriteBackgroundColours(RawRange, ColourDataProcessor.RgbColourToBgr(value));
        }

        public int[] FirstRowRgbColourZeroBased
        {
            get => RowAt(1).RgbBackgroundColoursZeroBased.SliceRow(0);
            set => RowAt(1).RgbBackgroundColoursZeroBased = value.To2DArray(ArrayOrientation.COLUMN);
        }

        public int[] FirstColumnRgbColourZeroBased
        {
            get => ColumnAt(1).RgbBackgroundColoursZeroBased.SliceColumn(0);
            set => ColumnAt(1).RgbBackgroundColoursZeroBased = value.To2DArray(ArrayOrientation.ROW);
        }

        public int Row => RawRange.Row;
        public int Column => RawRange.Column;

        public IWorksheetD Worksheet => DecoratorFactory.WorksheetD(RawRange.Worksheet);

        public IRangeD RowAt(int row)
        {
            return InteropTypeValidator.GetMapValidate<Range, IRangeD>
                (() => RawRange.Rows[row], DecoratorFactory.RangeD);
        }

        public IRangeD ColumnAt(int column) => InteropTypeValidator.MapValidate<Range, IRangeD>(RawRange.Columns[column],
            DecoratorFactory.RangeD);

        public void HideColumns()
        {
            RawRange.EntireColumn.Hidden = true;
        }

        public void HideRows()
        {
            RawRange.EntireRow.Hidden = true;
        }

        public void Clear()
        {
            RawRange.Clear();
        }

        public void ClearComments()
        {
            RawRange.ClearComments();
        }

        public void ClearData()
        {
            RawRange.Cells.ClearContents();
        }

        public void ClearDataAndHideColumns()
        {
            ClearData(); HideColumns();
        }

        public void ClearDataAndHideRows()
        {
            ClearData(); HideRows();
        }

        public void ReplaceWithValues()
        {
            ValueZeroBased = ValueZeroBased;
        }

        public bool IsOnSheet(IWorksheetD sourceSheet)
        {
            //non-mvp: define equality on worksheet decorators and use that here?
            return Worksheet.Name == sourceSheet.Name;
        }

        public void Delete()
        {
            try
            {
                RawRange.Delete();
            }
            catch (Exception e)
            {
                throw new InvalidOperationException("Could not delete range: " + AddressLocal, e);
            }
        }

        public void Insert()
        {
            RawRange.Insert();
        }

        public void Insert(XlInsertShiftDirection dir)
        {
            RawRange.Insert(dir);
        }

        [Obsolete("Do not use this Copy method unles strictly necessary - use the overloaded Copy method with a destination instead")]
        public void Copy()
        {
            RawRange.Copy();
        }

        public void Copy(IRangeD destination, XlPasteType xlPasteType)
        {
            try
            {
                RawRange.Copy();
                destination.RawRange.PasteSpecial(xlPasteType);
                RawRange.Application.CutCopyMode = 0;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(string.Format("Failed to copy data from {0} to {1}",
                    AddressLocal, destination?.AddressLocal), ex);
            }
        }

        public IList<IRangeD> AllAreas()
        {
            List<IRangeD> areaList = new List<IRangeD>();
            foreach (Range area in RawRange.Areas)
            {
                areaList.Add(DecoratorFactory.RangeD(area));
            }
            return areaList;
        }

        public IRangeD Cell(int row, int col)
        {            
            try
            {
                return InteropTypeValidator.GetMapValidate<Range, IRangeD>(
                () => RawRange.Cells[row, col], DecoratorFactory.RangeD);
            }
            catch (COMException e)
            {
                string address;
                try
                {
                    address = RawRange.AddressLocal;
                }
                catch (Exception)
                {
                    address = "<Unknown address>";
                }
                throw new InvalidOperationException(string.Format(
                    "Problem getting cell from range {0} with row, column coordinates ({1}, {2})",
                    address, row, col), e);
            }
        }

        public IList<IRangeD> Rows()
        {
            List<IRangeD> rowList = new List<IRangeD>();
            foreach (Range row in RawRange.Rows)
            {
                rowList.Add(DecoratorFactory.RangeD(row));
            }
            return rowList;
        }

        public IList<IRangeD> Columns()
        {
            List<IRangeD> columnList = new List<IRangeD>();
            foreach (Range col in RawRange.Columns)
            {
                columnList.Add(DecoratorFactory.RangeD(col));
            }
            return columnList;
        }

        public string? CommentTextOrNull => RawRange.Comment?.Text();

        public void AppendComment(string comment)
        {
            string? commentTextOrNull = CommentTextOrNull;
            if (commentTextOrNull == null)
            {
                AddComment(comment);
            }
            else
            {
                RawRange.Comment.Delete();
                AddComment(commentTextOrNull + comment);
            }
        }

        public void AddComment(string comment)
        {
            RawRange.AddComment(comment);
        }

        public IRangeD? VisibleRangeOrNull => CalculateVisibleRangeOrNull();

        private IRangeD? CalculateVisibleRangeOrNull() =>
            SpecialCellsOrNull(XlCellType.xlCellTypeVisible);

        //If there's an exception trying to get the special cells, this just returns null
        private IRangeD? SpecialCellsOrNull(XlCellType xlCellType) =>
            InteropTypeValidator.GetMapValidateOrNull<Range, IRangeD>(
                () => RawRange.SpecialCells(xlCellType), DecoratorFactory.RangeD);

        public void Fill(int colour)
        {
            SetColourForAllCells(colour);
        }

        private void SetColourForAllCells(int colour)
        {
            Cells.RawRange.Interior.Color = colour;
        }

        public void FillRgb(int rgbColour) => Fill(ColourDataProcessor.RgbColourToBgr(rgbColour));
        public int FirstCellColourBgr
        {
            get => FirstCell.Interior.ColourBgr;
            set => FirstCell.Interior.ColourBgr = value;
        }
        public int FirstCellRgbColour
        {
            get => ColourDataProcessor.BgrColourToRgb(FirstCellColourBgr);
            set => FirstCellColourBgr = ColourDataProcessor.RgbColourToBgr(value);
        }

        public int RowCount => RawRange.Rows.Count;
        public int ColumnCount => RawRange.Columns.Count;
        public int AreaCount => RawRange.Areas.Count;

        public string AddressLocal => RawRange.AddressLocal;
        public IWorksheetD Parent => InteropTypeValidator
            .GetMapValidate<Worksheet, IWorksheetD>(() => RawRange.Parent, DecoratorFactory.WorksheetD);

        public IRangeD this[int row, int column] => Cell(row, column);

        private IRangeD FirstCell => Cell(1, 1);

        public bool? HasFormula => GetHasFormula();

        private bool? GetHasFormula()
        {
            object hasFormulaRaw = RawRange.HasFormula;
            if (hasFormulaRaw == null) return null;
            if (hasFormulaRaw is bool hasFormula)
            {
                return hasFormula;
            }
            Log.Debug("Non-boolean returned by Interop layer for hasFormula. Interpreting this as false.");
            return false;
        }

        public IOneBasedArray<string> FirstColumnValue
        {
            get
            {
                return CsharpExtrasApi.NewOneBasedArray(FirstColumnValueZeroBased);
            }
            set
            {
                FirstColumnValueZeroBased = value.ZeroBasedEquivalent;
            }
        }

        public string? NameStrOrNull
        {
            set => RawRange.Name = value;
            get => NameOrNull?.Name;
        }

        public INameD? NameOrNull => InteropTypeValidator.GetMapValidateOrNull<Name, INameD>
            (() => RawRange.Name, DecoratorFactory.NameD);

        public IOneBasedArray<string> FirstRowValue
        {
            get => CsharpExtrasApi.NewOneBasedArray(FirstRowValueZeroBased);
            set => FirstRowValueZeroBased = value.ZeroBasedEquivalent;
        }

        public IOneBasedArray<string> FirstRowFormula
        {
            get => CsharpExtrasApi.NewOneBasedArray(FirstRowFormulaZeroBased);
            set => FirstRowFormulaZeroBased = value.ZeroBasedEquivalent;
        }

        public IOneBasedArray<string> FirstColumnFormula
        {
            get => CsharpExtrasApi.NewOneBasedArray(FirstColumnFormulaZeroBased);
            set => FirstColumnFormulaZeroBased = value.ZeroBasedEquivalent;
        }

        public int RowHeight { get => GetRowHeight(); set => RawRange.RowHeight = value; }
        private int GetRowHeight() => InteropTypeValidator.GetMapValidate<int, int>(() => RawRange.RowHeight,
            x => x);

        public int ColumnWidth { get => GetColumnWidth(); set => RawRange.ColumnWidth = value; }

        private int GetColumnWidth() => InteropTypeValidator.GetMapValidate<int, int>(() => RawRange.ColumnWidth,
            x => x);

        public bool WrapText { get => GetWrapText(); set => RawRange.WrapText = value; }

        private bool GetWrapText() => InteropTypeValidator.GetMapValidate<bool, bool>(() => RawRange.WrapText,
            x => x);

        public IEnumerable<IRangeD> AllCells => AllCellsEnumerate();

        private IEnumerable<IRangeD> AllCellsEnumerate()
        {
            foreach(object? cell in RawRange.Cells)
            {
                yield return InteropTypeValidator.MapValidate<Range, IRangeD>(cell, DecoratorFactory.RangeD);
            }
        }

        public IRangeD Cells => DecoratorFactory.RangeD(RawRange.Cells);

        public IRangeD SubArea(int startRow, int startColumn, int endRow, int endColumn)
        {
            IRangeD fromCell = Cell(startRow, startColumn);
            IRangeD toCell = Cell(endRow, endColumn);
            IRangeD newRange = Parent.Range(fromCell, toCell);
            return newRange;
        }

        public int LastRow => Row + RowCount - 1;

        public bool Hidden => InteropTypeValidator.GetMapValidate<bool, bool>(() => RawRange.Hidden, x=>x);

        public string? NumberFormatOrNull
        {
            get => RawRange.NumberFormat?.ToString();
            set => RawRange.NumberFormat = value;
        }

        public IRangeD MergeArea => DecoratorFactory.RangeD(RawRange.MergeArea);

        public IInteriorD Interior => DecoratorFactory.InteriorD(RawRange.Interior);

        public IOneBasedArray2D<string?> GetCellNames()
        {
            /*Non-mvp: Use performance-analysis to figure out if we need to scale this by some factor
             The current code assumes checking a name's refers-to range is about as expensive as checking
            a range's name*/
            int totalNamesInWorkbook = Worksheet.Parent.Names.Count();
            if (CellCount < totalNamesInWorkbook)
            {
                return GetCellNamesByCellsInRangeIteration();
            }
            return GetCellNamesByNamesInWorkbookIteration();
        }

        private IOneBasedArray2D<string?> GetCellNamesByCellsInRangeIteration()
        {
            //Non-mvp: Investigate is caching properties really necessary for a performance boost?
            int rowCount = RowCount;
            int columnCount = ColumnCount;
            IOneBasedArray2D<string?> names = CsharpExtrasApi.NewOneBasedArray2D<string?>(rowCount, columnCount);
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column < columnCount; column++)
                {
                    IRangeD cell = Cells[row, column];
                    names[row, column] = cell.NameStrOrNull;
                }
            }
            return names;
        }

        private IOneBasedArray2D<string?> GetCellNamesByNamesInWorkbookIteration()
        {
            IOneBasedArray2D<string?> names = CsharpExtrasApi.NewOneBasedArray2D<string?>(RowCount, ColumnCount);
            //Non-mvp: Investigate is caching properties really necessary for a performance boost?
            int startRow = Row;
            int startColumn = Column;
            int rowCount = RowCount;
            int endRow = Row + rowCount;
            int columnCount = ColumnCount;
            int endColumn = Column + columnCount;
            //Non-mvp: Investigate if this step is actually necessary - strings should default to null
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column < columnCount; column++)
                {
                    names[row, column] = null;
                }
            }
            INamesD allNamesInWorkbook = Worksheet.Parent.Names;
            foreach (INameD name in allNamesInWorkbook)
            {
                IRangeD? refersToRange = name.RefersToRangeOrNull;
                if (refersToRange != null)
                {
                    int absoluteRow = refersToRange.Row;
                    int absoluteColumn = refersToRange.Column;
                    if (refersToRange.IsSingleCell &&
                        IsWithinBounds(absoluteRow, absoluteColumn, startRow, startColumn, endRow, endColumn))
                    {
                        int relativeRow = absoluteRow - startRow + 1;
                        int relativeColumn = absoluteColumn - startColumn + 1;
                        names[relativeRow, relativeColumn] = name.Name;
                    }
                }
            }
            return names;
        }

        private bool IsWithinBounds(int row, int column, int startRow, int startColumn, int endRow, int endColumn)
        {
            return row >= startRow &&
                row <= endRow &&
                column >= startColumn &&
                column <= endColumn;
        }

        public void FilterToValues(int column, string[] values)
        {
            RawRange.AutoFilter(Field: column, Criteria1: values, Operator: XlAutoFilterOperator.xlFilterValues);
        }

        public void UnFilter(int column)
        {
            RawRange.AutoFilter(Field: column);
        }

        public IRangeD Offset(int row, int column)
        {
            return DecoratorFactory.RangeD(RawRange.Offset[row, column]);
        }

        public void TrySelect()
        {
            try
            {
                Select();
            }
            catch (Exception e)
            {
                Log.Warn("Select range failed. Swallowing exception.", e);
            }
        }

        public void Select()
        {
            RawRange.Worksheet.Select();
            RawRange.Select();
        }

        public string Address(bool absoluteColumn, bool absoluteRow)
        {
            return RawRange.Address[absoluteColumn, absoluteRow];
        }

        public bool IsFinite()
        {
            IRangeD entireRow = EntireRow;
            if (ColumnCount >= entireRow.ColumnCount) return false;
            IRangeD entireColumn = EntireColumn;
            if (RowCount >= entireColumn.RowCount) return false;
            return true;
        }

        public void TrySetRowHeight(int value)
        {
            try
            {
                RowHeight = value;
            }
            catch (Exception ex)
            {
                Log.Warn(string.Format("Failed to set row height of range '{0}' to {1}", AddressLocal, value), ex);
            }
        }

        public void TrySetColumnWidth(int value)
        {
            try
            {
                ColumnWidth = value;
            }
            catch (Exception ex)
            {
                Log.Warn(string.Format("Failed to set column width of range '{0}' to {1}", AddressLocal, value), ex);
            }
        }

        public void DoForAllAreas(Action<IRangeD> areaOperation)
        {
            foreach (IRangeD area in AllAreas())
            {
                areaOperation(area);
            }
        }

        //Non-mvp: consider overloading this or adding more optional parameters e.g. ignore blanks
        public void SetDropdownValidation(params string[] values)
        {
            string valuesJoined = string.Join(",", values);
            Microsoft.Office.Interop.Excel.Validation validation = RawRange.Validation;
            validation.Delete();
            validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertWarning, Formula1: valuesJoined, Formula2: Missing.Value);
            validation.InCellDropdown = true;
        }

        public bool AreSomeCellsMerged()
        {
            object mergeCells = RawRange.MergeCells;
            if (mergeCells is bool someCellsMerged)
            {
                if (someCellsMerged) return true;
            }
            return mergeCells == DBNull.Value;
        }

        public ISet<IRangeD> NonSingletonMergeAreas()
        {
            ISet<IRangeD> mergeAreas = ComputeAllMergeAreas();
            ISet<IRangeD> nonSingletonMergeAreas = new HashSet<IRangeD>();
            foreach (IRangeD range in mergeAreas)
            {
                if (range.CellCount != 1) nonSingletonMergeAreas.Add(range);
            }
            return nonSingletonMergeAreas;
        }
        private ISet<IRangeD> ComputeAllMergeAreas()
        {
            ISet<IRangeD> set = new HashSet<IRangeD>();
            foreach (IRangeD cell in AllCells)
            {
                IRangeD mergeArea = cell.MergeArea;
                set.Add(mergeArea);
            }
            return set;
        }

        public bool IsMergeSource()
        {
            if (!IsSingleCell) return false;
            ISet<IRangeD> mergeAreas = NonSingletonMergeAreas();
            foreach (IRangeD mergeArea in mergeAreas)
            {
                IRangeD topLeftCorner = mergeArea.Cell(1, 1);
                if (topLeftCorner.AddressLocal == AddressLocal)
                {
                    return true;
                }
            }
            return false;
        }

        public int GetNumberOfVisibleCells()
        {
            int cellCount = VisibleRangeOrNull != null ? VisibleRangeOrNull.CellCount : 0;
            return cellCount;
        }

        public bool AreAllCellsVisible() => GetNumberOfVisibleCells() == CellCount;
        public bool AreSomeCellsHidden() => !AreAllCellsVisible();

        public IRangeD ProjectRangeToSingleToColumn(int column)
        {
            IList<IRangeD> offsetAreas = AllAreas().Map(a => ProjectAreaToSingleColumn(a, column));
            if(offsetAreas.Count == 0)
            {
                throw new InvalidOperationException("Range unexpectedly has zero areas");
            }
            IRangeD unionRange = offsetAreas[0];
            foreach (IRangeD area in offsetAreas.Skip(1))
            {
                unionRange = Application.Union(unionRange, area);
            }
            return unionRange;
        }

        private IRangeD ProjectAreaToSingleColumn(IRangeD area, int column)
        {
            IRangeD resizedArea = area.Resize(area.RowCount, 1);
            IRangeD offsetArea = resizedArea.Offset(0, column - resizedArea.Column);
            return offsetArea;
        }

        public int LastUsedValueRow()
        {
            IOneBasedArray2D<string> values = Value;
            return values.LastUsedRow();
        }

        public int LastUsedFormulaRow()
        {
            IOneBasedArray2D<string> formulas = Formula;
            return formulas.LastUsedRow();
        }

        public int LastUsedRow()
        {
            int lastFormulaRow = LastUsedFormulaRow();
            int lastValueRow = LastUsedValueRow();
            return Math.Max(lastValueRow, lastFormulaRow);
        }

        public void ShowColumns()
        {
            RawRange.EntireColumn.Hidden = false;
        }

        public void ShowRows()
        {
            RawRange.EntireRow.Hidden = false;
        }

        public void CutInsert(IRangeD targetRange, XlInsertShiftDirection xlInsertShiftDirection)
        {
            try
            {
                IApplicationD targetApp = targetRange.Application;
                //Non-MVP: see if there's a better way to check that two ranges are in the same app
                if(Application.RawApplication != targetApp.RawApplication)
                {
                    throw new InvalidOperationException(
                        $"Cannot cut-insert range from a different application. " +
                        $"Attemped cut-insert from range {AddressLocal}" +
                        $" to range {targetRange.AddressLocal}");
                }
                RawRange.Cut();
                IWorksheetD targetWorksheet = targetRange.Worksheet;
                targetRange.Insert(xlInsertShiftDirection);
                targetApp.RawApplication.CutCopyMode = 0;
            }
            catch(Exception e)
            {
                throw new InvalidOperationException(
                    $"Exception thrown during cut-insert from range " +
                    $"{AddressLocal} to range {targetRange.AddressLocal}", e);
            }
        }

        public IRangeD Resize(int rowSize, int columnSize)
        {
            return DecoratorFactory.RangeD(RawRange.Resize[rowSize, columnSize]);
        }
    }
}
