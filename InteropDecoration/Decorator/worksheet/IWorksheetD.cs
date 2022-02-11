using InteropDecoration.Decorator.range;
using InteropDecoration.Decorator.listObjects;
using Microsoft.Office.Interop.Excel;
using InteropDecoration.Decorator.workbook;
using InteropDecoration.Decorator.comments;
using InteropDecoration.Decorator.tab;
using InteropDecoration.Decorator.application;
using CsharpExtras.Enumerable.OneBased;

namespace InteropDecoration.Decorator.worksheet
{
    public interface IWorksheetD
    {
        ITabD Tab { get; }
        IApplicationD Application { get; }

        //Non-mvp: return decorated type here instead
        IWorkbookD Parent { get; }
        Worksheet RawWorksheet { get; }

        ICommentsD Comments { get; }

        string Name { get; }

        bool IsVisible { get; }
        bool IsVeryHidden { get; }

        void HideSheet();
        void ShowSheet();

        /// <summary>
        /// Yields the smallest rectangle, starting at the upper lefmost cell A1, that contains the used range.
        /// </summary>
        IRangeD OriginUsedRange { get; }
        IRangeD UsedRange { get; }
        IListObjectsD ListObjectsD { get; }
        int OriginUsedRowsCount { get; }
        int OriginUsedColumnsCount { get; }
        IOneBasedArray2D<string> OriginUsedRangeValue { get; }
        IOneBasedArray2D<string> OriginUsedRangeFormula { get; }

        /// <summary>
        /// Return a RangeD for the ENTIRE row on the sheet.
        /// For performance reasons it is recomended to use OriginUsedRange.RowAt() instead.
        /// </summary>
        /// <param name="i">One based row index</param>
        IRangeD RowAt(int i);
        void DeleteRow(int row);

        /// <summary>
        /// Return a RangeD for the ENTIRE column on the sheet. Up to 1 million cells in size.
        /// For performance reasons it is recomended to use OriginUsedRange.ColumnAt() instead.
        /// </summary>
        /// <param name="i">One based column index</param>
        IRangeD ColumnAt(int i);

        IRangeD ColumnOriginUsedRange(int column);
        IRangeD RowOriginUsedRange(int row);

        /// <summary>Gets the range with the same address on this sheet</summary>
        IRangeD GetCorrespondingRange(IRangeD range);
        IRangeD Range(IRangeD from, IRangeD to);
        IRangeD Range(int fromRow, int fromCol, int toRow, int toCol);
        IRangeD Cells(int row, int column);
        bool IsRowBlank(int row);
        bool IsColumnBlank(int col);
        void Activate();
        void FreezeUpToRow(int rowIndex);

        /// <summary>
        /// Removes all rows in the given range that match the given predicates. The predicates are joined using AND logic.
        /// </summary>
        /// <param name="columnMatchPredicates">Each predicate must match on the associated column in order for the row to be deleted.</param>        
        /// <returns>A collection of values for the deleted rows. Each value is an array of strings of values on that row, up to the used range.</returns>
        ICollection<string[]> RemoveAllRowsWhere(Dictionary<int, Func<string, bool>> columnMatchPredicates, int fromRow, int toRow);

        /// <summary>Similar to RemoveAllRowsWhere</summary>        
        ICollection<string[]> RemoveAllColumnsWhere(Dictionary<int, Func<string, bool>> rowMatchPredicates, int fromColumn, int toColumn);
        
        ICollection<string[]> BlankAndHideAllRowsWhere(int column, Func<string, bool> matcher, int fromRow, int toRow);
        ICollection<string[]> BlankAndHideAllColumnsWhere(int row, Func<string, bool> matcher, int fromColumn, int toColumn);

        ICollection<string[]> RemoveAllRowsWhere(int column, Func<string, bool> matcher, int fromRow, int toRow);
        ICollection<string[]> RemoveAllColumnsWhere(int row, Func<string, bool> matcher, int fromColumn, int toColumn);
        
        /// <summary>Replaces all the worksheets cells with values, leaving no formulas.</summary>
        void ReplaceWithValues();
        void TryAddTableFilter(string tableName, int column, string[] constraintArray);
        void TryRemoveTableFilter(string tableName, int column);
        bool TryClearAllFilters();

        void ClearAll();    
        void ClearUsedRange();

        void UnhideColumns();
        void UnhideRows();

        /// <returns>If there are any visible cells then a non-empty range containing those cells is returned.
        /// Otherwise, null is returned.</returns>
        IRangeD? GetVisibleRangeOrNull();

        IRangeD? FindCellsWithFormulaErrorsOrNull();

        int CountCellsWithConstantErrors();
        IRangeD? FindCellsWithConstantErrorsOrNull();
        IRangeD CroppedOriginUsedRange(int maxRow);

        void RefreshAllPivotTables();
        void TryRefreshAllPivotTables();
        IOneBasedArray<string> ValueRow(int row);
        void WriteValueRow(int row, IOneBasedArray<string> valueRow);
        IOneBasedArray<string> FormulaRow(int row);
        void WriteFormulaRow(int row, IOneBasedArray<string> formulaRow);
        IOneBasedArray<string> ValueColumn(int Column);
        void WriteValueColumn(int Column, IOneBasedArray<string> valueColumn);
        IOneBasedArray<string> FormulaColumn(int Column);
        void WriteFormulaColumn(int Column, IOneBasedArray<string> formulaColumn);
    }
}
