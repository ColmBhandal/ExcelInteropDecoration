namespace ExcelInteropDecoration.Decorator.util
{
    public interface IInteropStringProcessor
    {
        (string sheetName, string relativeAddress) SplitAddress(string absoluteAddress);
        //non-mvp: Could also implement this for just sheet name and address
        //non-mvp: Test if excel throwns an exception if the workbook name doesn't have an extension. Right now it assumes an extension.
        string CombineAddress(string workbookNameWithExtension, string sheetName, string address);

        string BuildAddress(string workbookNameWithExtension, string sheetName, int row, int column);
        string BuildAddress(string sheetName, int row, int column);

        string IntToAlphabet(int i);
        int AlphabetToInt(string alfa);

        string RowColToLocalAddress(int row, int col);
        
        (int? row, int? col) LocalAddressToRowColPair(string localAddress);
        (string fromCellAddress, string toCellAddress) SplitAreaAddressIntoCorners(string areaAddress);
    }
}