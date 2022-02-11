using CsharpExtras.Extensions;
using ExcelInteropDecoration._base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelInteropDecoration.Decorator.util
{
    class InteropStringProcessorImpl : BaseClass, IInteropStringProcessor
    {
        private const int AsciiAlphabetStartIndex = 64;

        private const string ExcelCellAddressRegexLax = "(?<absCol>\\$?)(?<colRef>[A-Z]*)(?<absRow>\\$?)(?<rowRef>[0-9]*)";
        //Non-mvp: Change * to + and add functionality for ignoring matches for things inside quotes
        private const string ExcelAreaAddressRegex = "(?<sheetName>'?[\\w\\-]*'?!)?(?<ref1>[A-Z0-9$]+):?(?<ref2>[A-Z0-9$]+)?";

        public InteropStringProcessorImpl(IInteropDAPI interopDApi) : base(interopDApi)
        {
        }

        public (string sheetName, string relativeAddress) SplitAddress(string absoluteAddress)
        {
            string[] splitStr = absoluteAddress.Split('!');
            if (splitStr.Length != 2)
            {
                throw new ArgumentException(string.Format(
                    "Cannot parse {0} as an absolute address. Format should be two strings separated by '!'.", absoluteAddress));
            }
            //Else, we know we have two strings. Return them.
            string sheetNameNoApostrophes = removeApostrophes(splitStr[0]);
            string relativeAddressNoApostrophes = removeApostrophes(splitStr[1]);
            return ((sheetNameNoApostrophes, relativeAddressNoApostrophes));
        }

        private string removeApostrophes(string s)
        {
            return s.RemoveRegexMatches("'");
        }

        public string CombineAddress(string workbookNameWithExtension, string sheetName, string localAddress)
        {
            string wbNameNoApostrophes = removeApostrophes(workbookNameWithExtension);
            string sheetNameNoApostrophes = removeApostrophes(sheetName);
            string localAddressNoApostrpohes = removeApostrophes(localAddress);
            return string.Format("'[{0}]{1}'!{2}", wbNameNoApostrophes, sheetNameNoApostrophes, localAddressNoApostrpohes);
        }

        public string BuildAddress(string workbookNameWithExtension, string sheetName, int row, int column)
        {
            string columnStr = IntToAlphabet(column);
            string sheetPrefix = $"'{sheetName}'";

            if (!string.IsNullOrWhiteSpace(workbookNameWithExtension))
            {
                sheetPrefix = $"'[{workbookNameWithExtension}]{sheetName}'";
            }

            return $"{sheetPrefix}!${columnStr}${row}";
        }

        public string BuildAddress(string sheetName, int row, int column)
        {
            return BuildAddress("", sheetName, row, column);
        }

        public string IntToAlphabet(int i)
        {
            if (i < 1)
            {
                throw new InvalidOperationException("Integer value must be at least 1. Ran with: " + i);
            }

            int dividend = i;
            string stringResult = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                stringResult = (char)(AsciiAlphabetStartIndex + 1 + modulo) + stringResult;
                dividend = (dividend - modulo) / 26;
            }

            return stringResult;
        }

        public int AlphabetToInt(string alfa)
        {
            if (string.IsNullOrWhiteSpace(alfa))
            {
                throw new InvalidOperationException("String value must not be null or whitespace");
            }

            char[] letters = alfa.ToArray();
            int result = 0;

            for (int i = 0; i < letters.Length; i++)
            {
                char c = letters[i];
                result *= 26;
                result += LetterToInt(c);
            }

            return result;
        }

        private int LetterToInt(char letter)
        {
            return letter - AsciiAlphabetStartIndex;
        }

        public string RowColToLocalAddress(int row, int col)
        {
            return IntToAlphabet(col) + row;
        }

        /// <summary>
        /// Given an addresss, which might span over an area bigger than one cell, this splits it into the two corners defining the area.
        /// </summary>
        /// <param name="possiblyRectangularAddress">The address of either a single cell or an area.</param>
        /// <returns>A pair of addresses representing single cells which define the area as its two corners.</returns>
        public (string fromCellAddress, string toCellAddress) SplitAreaAddressIntoCorners(string areaAddress)
        {
            Regex addressRegex = new Regex(ExcelAreaAddressRegex);
            Match match = addressRegex.Match(areaAddress);
            string fromAddress = match.Groups["ref1"].Value;
            string toAddress = match.Groups["ref2"].Value;
            if (string.IsNullOrWhiteSpace(toAddress))
            {
                toAddress = fromAddress;
            }
            return (fromAddress, toAddress);
        }

        public (int? row, int? col) LocalAddressToRowColPair(string localAddress)
        {
            Regex addressRegex = new Regex(ExcelCellAddressRegexLax);
            Match match = addressRegex.Match(localAddress);
            string column = match.Groups["colRef"].Value;
            string rowStr = match.Groups["rowRef"].Value;
            int? col = null;
            if(!string.IsNullOrWhiteSpace(column)) col = AlphabetToInt(column);
            int? row = null;
            if(int.TryParse(rowStr, out int rowInt))
            {
                row = rowInt;
            }
            return (row, col);
        }
    }
}
