using ExcelInteropDecoration.Decorator.worksheet;
using Microsoft.Office.Interop.Excel;
using CsharpExtras.Enumerable.OneBased;
using ExcelInteropDecoration.Decorator.application;
using ExcelInteropDecoration.Decorator.comment;
using Range = Microsoft.Office.Interop.Excel.Range;
using ExcelInteropDecoration.Decorator.interior;

namespace ExcelInteropDecoration.Decorator.range
{
    public interface IRangeD
    {
        int RowCount { get; }
        int ColumnCount { get; }

        IRangeD MergeArea { get; }
        void DoForAllAreas(Action<IRangeD> areaOperation);
        IWorksheetD Worksheet { get; }
        IApplicationD Application { get; }
        Range RawRange { get; }

        bool IsOnSheet(IWorksheetD sourceSheet);
        bool IsFinite();

        IRangeD this[int row, int column] { get; }
        int AreaCount { get; }

        void Fill(int colour);
        
        /// <summary>
        /// If there are any visible cells then returns a range representing these cells. Otherwise returns null.
        /// </summary>
        IRangeD? VisibleRangeOrNull { get; }

        void CutInsert(IRangeD targetRange, XlInsertShiftDirection xlInsertShiftDirection);

        IList<IRangeD> AllAreas();
        void Clear();
        bool AreSomeCellsHidden();
        IRangeD Cell(int row, int col);
        IRangeD Cells { get; }
        IRangeD SubArea(int startRow, int startColumn, int endRow, int endColumn);

        IList<IRangeD> Rows();
        IList<IRangeD> Columns();

        IOneBasedArray2D<string> Value { get; set; }
        IOneBasedArray2D<string> Formula { get; set; }
        string[,] ValueZeroBased { get; set; }
        int[,] RgbBackgroundColoursZeroBased { get; set; }

        string FirstCellValue { get; set; }
        string FirstCellFormula { get; set; }
        int FirstCellColourBgr { get; set; }
        int FirstCellRgbColour { get; set; }
        
        string[] FirstRowValueZeroBased { get; set; }
        string[] FirstRowFormulaZeroBased { get; set; }
        int[] FirstRowRgbColourZeroBased { get; set; }

        //non-mvp: make this a One-based array
        string[] FirstColumnValueZeroBased { get; set; }

        IOneBasedArray<string> FirstColumnValue { get; set; }
        IRangeD Offset(int row, int column);

        string[] FirstColumnFormulaZeroBased { get; set; }
        int[] FirstColumnRgbColourZeroBased { get; set; }
        IOneBasedArray<int> FirstColumnRgbColour { get; set; }     

        /// <summary>
        /// Get the sub-range defined at the given row. 
        /// The returned range will include only the columns included in the source range.
        /// </summary>
        /// <param name="row">One based row index</param>
        IRangeD RowAt(int row);
        void SetDropdownValidation(params string[] values);

        /// <summary>
        /// Get the sub-range defined at the given column. 
        /// The returned range will include only the rows included in the source range.
        /// </summary>
        /// <param name="row">One based column index</param>
        IRangeD ColumnAt(int column);

        string Address(bool absoluteColumn, bool absoluteRow);
        string AddressLocal { get; }
        string? NameStrOrNull { get;  set; }

        /// <summary>
        /// Gets all the names at each cell in this range. Cells without a name will have a null entry in the array.
        /// </summary>
        IOneBasedArray2D<string?> GetCellNames();
        
        void HideRows();
        void HideColumns();
        void ShowColumns();
        void ShowRows();
        void Select();
        //Non-mvp: rename to clear contents (so as to match the underlying function)
        void ClearData();
        void ClearDataAndHideColumns();
        void ClearDataAndHideRows();

        void Insert();
        void Insert(XlInsertShiftDirection dir);

        void Copy();
        void Copy(IRangeD destination, XlPasteType xlPasteType);

        /// <summary>
        /// Replaces all cells with their values at the time of the call. No more formulas will exist after this call.
        /// </summary>
        void ReplaceWithValues();
        IWorksheetD Parent { get; }
        IOneBasedArray<string> FirstRowValue { get; set; }

        bool AreSomeCellsMerged();

        IOneBasedArray<string> FirstRowFormula { get; set; }
        IOneBasedArray<string> FirstColumnFormula { get; set; }      
        int CellCount { get; }
        int RowHeight { get; set; }
        void TrySetRowHeight(int value);
        int ColumnWidth { get; set; }
        void TrySetColumnWidth(int value);

        /// <returns>True iff this cell is the top left corner of a merged area.</returns>
        bool IsMergeSource();
        ISet<IRangeD> NonSingletonMergeAreas();

        IEnumerable<IRangeD> AllCells { get; }
        IRangeD EntireRow { get; }
        IRangeD EntireColumn { get; }
        int LastRow { get; }

        /// <returns>The last row in the range that contains a non-empty string in one of its values</returns>
        int LastUsedValueRow();

        /// <returns>The last row in the range that contains a non-empty string in one of its formulas</returns>
        int LastUsedFormulaRow();

        /// <returns>The last row in the range that contains a non-empty string in one of its values or formulas</returns>
        int LastUsedRow();
        bool Hidden { get; }
        string? NumberFormatOrNull { get; set; }
        IOneBasedArray2D<int> RgbBackgroundColours { get; set; }
        bool IsSingleCell { get; }
        bool AreAllCellsVisible();

        string? CommentTextOrNull { get; }
        bool? HasFormula { get; }

        void ClearComments();
        void AddComment(string comment);

        void Delete();
        void FilterToValues(int column, string[] values);
        void UnFilter(int column);
        void TrySelect();
        int GetNumberOfVisibleCells();

        /// <summary>
        /// Project the given range to a single column. The column can either be inside or outside the given range.
        /// For example, if the current range is A1:C5, when this method is called with column 4, the returned range will be D1:D5.
        /// </summary>
        IRangeD ProjectRangeToSingleToColumn(int column);
        void AppendComment(string comment);        

        bool WrapText { get; set; }
        ICommentD? CommentOrNull { get; }
        string[,] FormulaZeroBased { get; set; }
        IInteriorD Interior { get; }
        int Row { get; }
        int Column { get; }

        IRangeD Resize(int rows, int columns);
    }
}
