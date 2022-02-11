using ExcelInteropDecoration.Decorator.range;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using CsharpExtras.Enumerable.OneBased;
using ExcelInteropDecoration.Decorator.listObjects;
using ExcelInteropDecoration.Decorator.workbook;
using PerformanceRecorder.Attribute;
using ExcelInteropDecoration.Decorator.comments;
using ExcelInteropDecoration.Decorator.tab;
using ExcelInteropDecoration.Decorator.application;
using ExcelInteropDecoration.Decorator._base;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelInteropDecoration.Decorator.worksheet
{
    class WorksheetDImpl : DecoratorBase, IWorksheetD
    {
        public IWorkbookD Parent => InteropTypeValidator.GetMapValidate<Workbook, IWorkbookD>
            (() => RawWorksheet.Parent, DecoratorFactory.WorkbookD);

        [Obsolete("Don't use raw Excel Interop unless strictly necessary - use decoration layer instead")]
        public Worksheet RawWorksheet { get; }

        public virtual string Name => RawWorksheet.Name;

        public bool IsVisible => RawWorksheet.Visible == XlSheetVisibility.xlSheetVisible;

        public bool IsVeryHidden => RawWorksheet.Visible == XlSheetVisibility.xlSheetVeryHidden;

        public void HideSheet() => RawWorksheet.Visible = XlSheetVisibility.xlSheetHidden;

        public void ShowSheet() => RawWorksheet.Visible = XlSheetVisibility.xlSheetVisible;

        public WorksheetDImpl(IInteropDAPI api, Worksheet rawWorksheet) : base(api)
        {
            RawWorksheet = rawWorksheet ?? throw new ArgumentNullException(nameof(rawWorksheet));
        }
        public int OriginUsedRowsCount => OriginUsedRange.RowCount;
        public int OriginUsedColumnsCount => OriginUsedRange.ColumnCount;

        public IOneBasedArray<string> ValueRow(int row) => OriginUsedRange.RowAt(row).FirstRowValue;
        public void WriteValueRow(int row, IOneBasedArray<string> valueRow)
        {
            OriginUsedRange.RowAt(row).FirstRowValue = valueRow;
        }

        public IOneBasedArray<string> FormulaRow(int row) => OriginUsedRange.RowAt(row).FirstRowFormula;
        public void WriteFormulaRow(int row, IOneBasedArray<string> formulaRow)
        {
            OriginUsedRange.RowAt(row).FirstRowFormula = formulaRow;
        }


        public IOneBasedArray<string> ValueColumn(int Column) => OriginUsedRange.ColumnAt(Column).FirstColumnValue;
        public void WriteValueColumn(int Column, IOneBasedArray<string> valueColumn)
        {
            OriginUsedRange.ColumnAt(Column).FirstColumnValue = valueColumn;
        }

        public IOneBasedArray<string> FormulaColumn(int Column) => OriginUsedRange.ColumnAt(Column).FirstColumnFormula;
        public void WriteFormulaColumn(int Column, IOneBasedArray<string> formulaColumn)
        {
            OriginUsedRange.ColumnAt(Column).FirstColumnFormula = formulaColumn;
        }

        public IOneBasedArray2D<string> OriginUsedRangeValue => OriginUsedRange.Value;
        public IOneBasedArray2D<string> OriginUsedRangeFormula => OriginUsedRange.Formula;

        public IRangeD OriginUsedRange
        {
            get
            {
                Range usedRange = RawWorksheet.UsedRange;
                int rowsBefore = usedRange.Row - 1;
                int rowBound = rowsBefore + usedRange.Rows.Count;
                int columnsBefore = usedRange.Column - 1;
                int columnBound = columnsBefore + usedRange.Columns.Count;
                Range result = RawWorksheet.Range[RawWorksheet.Cells[1, 1], RawWorksheet.Cells[rowBound, columnBound]];
                return DecoratorFactory.RangeD(result);
            }
        }
        public IRangeD UsedRange => DecoratorFactory.RangeD(RawWorksheet.UsedRange);

        public IListObjectsD ListObjectsD => DecoratorFactory.ListObjectsD(RawWorksheet.ListObjects);

        public ICommentsD Comments => DecoratorFactory.CommentsD(RawWorksheet.Comments);

        public ITabD Tab => DecoratorFactory.TabD(RawWorksheet.Tab);

        public IApplicationD Application => Parent.Application;

        public IRangeD GetCorrespondingRange(IRangeD sourceRange) =>
            DecoratorFactory.RangeD(RawWorksheet.Range[sourceRange.AddressLocal]);

        public IRangeD RowAt(int row) => InteropTypeValidator.GetMapValidate<Range, IRangeD>
            (() => RawWorksheet.Rows[row], DecoratorFactory.RangeD);

        public IRangeD ColumnAt(int column) => InteropTypeValidator.GetMapValidate<Range, IRangeD>
            (() => RawWorksheet.Columns[column], DecoratorFactory.RangeD);

        public IRangeD ColumnOriginUsedRange(int column)
        {
            return Range(1, column, OriginUsedRowsCount, column);
        }
        public IRangeD RowOriginUsedRange(int row)
        {
            return Range(row, 1, row, OriginUsedColumnsCount);
        }
        
        public bool IsRowBlank(int row)
        {
            string[] rowVal = OriginUsedRange.RowAt(row).FirstRowValueZeroBased;
            foreach(string str in rowVal)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    return false;
                }
            }
            return true;
        }
        public bool IsColumnBlank(int col)
        {
            string[] colVal = OriginUsedRange.ColumnAt(col).FirstColumnValueZeroBased;
            foreach (string str in colVal)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    return false;
                }
            }
            return true;
        }

        public IRangeD Cells(int row, int column)
        {
            try
            {
                return InteropTypeValidator.GetMapValidate<Range, IRangeD>
                    (() => RawWorksheet.Cells[row, column], DecoratorFactory.RangeD);
            }
            catch (Exception ex)
            {
                string msg = string.Format(
                    "Problem getting cell ({0}, {1}) on sheet {2}",
                    row, column, RawWorksheet.Name);
                Log.Error(msg, ex);
                throw new InvalidOperationException(msg, ex);
            }
        }

        public IRangeD Range(string address) => DecoratorFactory.RangeD(RawWorksheet.Range[address]);
        public IRangeD Range(IRangeD from, IRangeD to)
        {
            return DecoratorFactory.RangeD(RawWorksheet.Range[from.RawRange, to.RawRange]);
        }
        public IRangeD Range(int fromRow, int fromCol, int toRow, int toCol) =>
            Range(Cells(fromRow, fromCol), Cells(toRow, toCol));


        //TODO: Test this on Themes sheet to see if it works for table filters
        /// <summary>
        /// Tries to clear both table filters and regular sheet filters, without failing on exceptions.
        /// </summary>
        /// <returns>True iff there were no exceptions encountered.</returns>
        public bool TryClearAllFilters()
        {
            bool status = true;
            try
            {
                if (RawWorksheet.AutoFilter != null)
                {
                    RawWorksheet.AutoFilter.ShowAllData();
                }
            }
            catch(Exception ex)
            {
                Log.Error(string.Format("Clear filters encountered an issue. Sheet: {0}", Name), ex);
                status = false;
            }
            try
            {
                ListObjects listObjects = RawWorksheet.ListObjects;                
                foreach(ListObject listObject in listObjects)
                {
                    listObject.AutoFilter.ShowAllData();
                }
            }
            catch (Exception ex)
            {
                Log.Error(string.Format("Clear table filters encountered an issue. Sheet: {0}", Name), ex);
                status = false;
            }
            return status;
        }

        public void TryAddTableFilter(string tableName, int column, string[] values)
        {
            TryInvokeTableRangeActionIfTableExists(tableName,
                range => range.FilterToValues(column, values), "Add-Table-Filter");
        }

        public void TryRemoveTableFilter(string tableName, int column)
        {
            TryInvokeTableRangeActionIfTableExists(tableName,
                range => range.UnFilter(column), "Remove-Table-Filter");
        }

        private void TryInvokeTableRangeActionIfTableExists(string tableName, Action<IRangeD> rangeOperator, string actionName)
        {
            IRangeD? tableRange = TableRangeOrNull(tableName);
            if (tableRange == null)
            {
                Log.Warn($"Did not perform action {actionName} for table " +
                    $"{tableName} on sheet {Name}. The table does not appear to exist.");
                return;
            }
            rangeOperator(tableRange);
        }

        private IRangeD? TableRangeOrNull(string tableName)
        {
            IListObjectsD listObjects = ListObjectsD;
            if (!listObjects.HasTable(tableName))
            {
                return null;
            }
            IListObjectD table = listObjects[tableName];
            return table.Range;
        }

        public ICollection<string[]> RemoveAllRowsWhere(int column, Func<string, bool> matcher, int fromRow, int toRow)
        {
            Dictionary<int, Func<string, bool>> dict = new Dictionary<int, Func<string, bool>>();
            dict.Add(column, matcher);
            return RemoveAllRowsWhere(dict, fromRow, toRow);
        }

        public ICollection<string[]> RemoveAllColumnsWhere(int row, Func<string, bool> matcher, int fromColumn, int toColumn)
        {
            Dictionary<int, Func<string, bool>> dict = new Dictionary<int, Func<string, bool>>();
            dict.Add(row, matcher);
            return RemoveAllColumnsWhere(dict, fromColumn, toColumn);
        }

        public ICollection<string[]> BlankAndHideAllRowsWhere(int column, Func<string, bool> matcher, int fromRow, int toRow)
        {
            Dictionary<int, Func<string, bool>> dict = new Dictionary<int, Func<string, bool>>();
            dict.Add(column, matcher);
            return BlankAndHideAllRowsWhere(dict, fromRow, toRow);
        }

        public ICollection<string[]> BlankAndHideAllColumnsWhere(int row, Func<string, bool> matcher, int fromColumn, int toColumn)
        {
            Dictionary<int, Func<string, bool>> dict = new Dictionary<int, Func<string, bool>>();
            dict.Add(row, matcher);
            return BlankAndHideAllColumnsWhere(dict, fromColumn, toColumn);
        }

        private ICollection<string[]> BlankAndHideAllColumnsWhere(Dictionary<int, Func<string, bool>> dict, int fromColumn, int toColumn)
        {
            return ProcessAllColumnsWhere(dict, fromColumn, toColumn, BlankAndHideColumnAndGetValues);
        }

        private ICollection<string[]> BlankAndHideAllRowsWhere(Dictionary<int, Func<string, bool>> dict, int fromRow, int toRow)
        {
            return ProcessAllRowsWhere(dict, fromRow, toRow, BlankAndHideRowAndGetValues);
        }
        private string[] BlankAndHideColumnAndGetValues(IRangeD range)
        {
            return ProcessRangeAndGetValues(range, ClearAndHideColumns);
        }

        private string[] BlankAndHideRowAndGetValues(IRangeD range)
        {
            return ProcessRangeAndGetValues(range, ClearAndHideRows);
        }

        private void ClearAndHideColumns(IRangeD range)
        {
            range.ClearData(); range.HideColumns();
        }

        private void ClearAndHideRows(IRangeD range)
        {
            range.ClearData(); range.HideRows();
        }

        //non-mvp: refactor to use a query object which selects and then deletes based on the selection
        public ICollection<string[]> RemoveAllRowsWhere(Dictionary<int, Func<string, bool>> columnMatchPredicates, int fromRow, int toRow)
        {
            return ProcessAllRowsWhere(columnMatchPredicates, fromRow, toRow, DeleteRangeAndGetValues); ;
        }

        //non-mvp: refactor to use a query object which selects and then deletes based on the selection
        public ICollection<string[]> RemoveAllColumnsWhere(Dictionary<int, Func<string, bool>> rowMatchPredicates, int fromColumn, int toColumn)
        {
            return ProcessAllColumnsWhere(rowMatchPredicates, fromColumn, toColumn, DeleteRangeAndGetValues);
        }

        private string[] DeleteRangeAndGetValues(IRangeD range)
        {            
            return ProcessRangeAndGetValues(range, r => r.Delete());
        }

        private string[] ProcessRangeAndGetValues(IRangeD range, Action<IRangeD> process)
        {
            string[] vals = range.FirstColumnValueZeroBased;
            process(range);
            return vals;
        }

        //non-mvp: refactor to use a query object which selects and then deletes based on the selection
        public ICollection<T> ProcessAllRowsWhere<T>(Dictionary<int, Func<string, bool>> rowMatchPredicates, int fromRow, int toRow,
            Func<IRangeD, T> processor)
        {
            //swap so we always have toRow <= fromRow
            if (fromRow < toRow)
            {
                int temp = fromRow;
                fromRow = toRow;
                toRow = temp;
            }
            ICollection<T> coll = new HashSet<T>();

            //Loop backwards so that row indices don't change after deletions
            for (int row = fromRow; row > toRow; row--)
            {
                IRangeD colRange = RowAt(row);
                string[] vals = colRange.FirstRowValueZeroBased;
                if (Matches(rowMatchPredicates, vals))
                {
                    coll.Add(processor(colRange));
                }
            }
            return coll;
        }

        //non-mvp: refactor to use a query object which selects and then deletes based on the selection
        public ICollection<T> ProcessAllColumnsWhere<T>(Dictionary<int, Func<string, bool>> rowMatchPredicates, int fromColumn, int toColumn,
            Func<IRangeD, T> processor)
        {
            //swap so we always have toColumn <= fromColumn
            if (fromColumn < toColumn)
            {
                int temp = fromColumn;
                fromColumn = toColumn;
                toColumn = temp;
            }
            ICollection<T> coll = new HashSet<T>();

            //Loop backwards so that column indices don't change after deletions
            for (int column = fromColumn; column > toColumn; column--)
            {
                IRangeD colRange = ColumnAt(column);
                string[] vals = colRange.FirstColumnValueZeroBased;
                if (Matches(rowMatchPredicates, vals))
                {                                        
                    coll.Add(processor(colRange));
                }
            }
            return coll;
        }

        private bool Matches(Dictionary<int, Func<string, bool>> predicates, string[] values)
        {
            foreach(int key in predicates.Keys)
            {
                //-1 because we're working 0-based in C# array, but predicate is 1-based as it works off Excel sheet
                if (!predicates[key](values[key -1])) return false;
            }
            return true;
        }

        public void ReplaceWithValues()
        {
            OriginUsedRange.ReplaceWithValues();
        }

        public void ClearAll()
        {
            RawWorksheet.Cells.Clear();
        }

        public void ClearUsedRange()
        {
            RawWorksheet.UsedRange.Clear();
        }

        public void UnhideColumns()
        {
            RawWorksheet.Columns.Hidden = false;
        }

        public void UnhideRows()
        {
            RawWorksheet.Rows.Hidden = false;
        }

        public IRangeD? GetVisibleRangeOrNull() =>
            OriginUsedRange.VisibleRangeOrNull;

        public void Activate()
        {
            RawWorksheet.Activate();
        }

        public void FreezeUpToRow(int rowIndex)
        {
            Activate();
            Window activeWindow = Parent.Application.ActiveWindow;
            activeWindow.SplitRow = rowIndex;
            activeWindow.FreezePanes = true;
        }

        [PerformanceLogging]
        public IRangeD? FindCellsWithFormulaErrorsOrNull() =>
            FindCellsOrNull(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlErrors);

        [PerformanceLogging]
        public IRangeD? FindCellsWithConstantErrorsOrNull() =>
            FindCellsOrNull(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlErrors);

        [PerformanceLogging]
        private IRangeD? FindCellsOrNull(XlCellType cellType, XlSpecialCellsValue valueType)
        {
            Range? rawCells = FindCellsOrNullRaw(cellType, valueType);
            if (rawCells != null)
            {
                return DecoratorFactory.RangeD(rawCells);
            }
            return null;
        }

        [PerformanceLogging]
        private Range? FindCellsOrNullRaw(XlCellType cellType, XlSpecialCellsValue valueType)
        {
            try
            {
                return RawWorksheet.Cells.SpecialCells(cellType, valueType);
            }
            catch
            {
                return null;
            }
        }

        public int CountCellsWithConstantErrors()
        {
            IRangeD? cellsWithConstantErrors = FindCellsWithConstantErrorsOrNull();
            if(cellsWithConstantErrors != null)
            {
                return cellsWithConstantErrors.CellCount;
            }
            return 0;
        }

        public void DeleteRow(int row)
        {
            RowAt(row).Delete();
        }

        public IRangeD CroppedOriginUsedRange(int maxRow)
        {
            int columnCount = OriginUsedRange.ColumnCount;
            return Range(1, 1, maxRow, columnCount);
        }

        public IOneBasedArray2D<string> CroppedOriginUsedRangeValues(int maxRow)
        {
            IRangeD croppedUsedRange = CroppedOriginUsedRange(maxRow);
            return croppedUsedRange.Value;
        }

        public IOneBasedArray2D<string> CroppedOriginUsedRangeFormulas(int maxRow)
        {
            IRangeD croppedUsedRange = CroppedOriginUsedRange(maxRow);
            return croppedUsedRange.Formula;
        }

        public void RefreshAllPivotTables()
        {
            // non-mvp: Add decorators for pivot tables
            PivotTables pivotTables;
            try
            {
                pivotTables = (PivotTables)RawWorksheet.PivotTables();
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Failed to retrieve pivot tables from worksheet. Unable to cast to expected type.", ex);
            }

            try
            {
                foreach (object pivotTable in pivotTables)
                {
                    ((PivotTable)pivotTable).RefreshTable();
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Failed to refresh pivot table. Unable to cast to expected type.", ex);
            }
        }

        public void TryRefreshAllPivotTables()
        {
            try
            {
                RefreshAllPivotTables();
            }
            catch (Exception ex)
            {
                Log.Error("Error caught while refreshing pivot tables", ex);
            }
        }
    }
}
