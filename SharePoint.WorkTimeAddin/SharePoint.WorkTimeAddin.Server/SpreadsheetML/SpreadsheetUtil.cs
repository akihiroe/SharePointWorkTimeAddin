using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePoint.WorkTimeAddin.SpreadsheetML
{
    /// <summary>
    /// Spreadsheetでよく利用する処理を提供します
    /// </summary>
    internal static class SpreadsheetUtil
    {
        /// <summary>
        /// 列文字列を列番号に変換します。
        /// </summary>
        /// <param name="columnLetter">列文字列</param>
        /// <returns>列番号</returns>
        public static int ConvertColumnIndex(string columnLetter)   
        {
            //性能対策
            int columnIndex = 0;
            foreach (char c in columnLetter)
            {
                columnIndex = columnIndex * 26 + ((int)c) - 64;
            }
            return columnIndex;
            //LINQ版 A-Zを26進数と見立て変換
            //return columnLetter.ToCharArray().Select(x => (int)x - 64).Aggregate((x, y) => x * 26 + y);
        }

        /// <summary>
        /// 列番号を列文字列に変換します。
        /// </summary>
        /// <param name="columnIndex"列番号></param>
        /// <returns>列文字列</returns>
        public static string ConvertColumnLetter(int columnIndex)
        {
            int idx = columnIndex;
            string columnLetter = "";
            while (idx > 0)
            {
                int modulo = (idx - 1) % 26;
                columnLetter = Convert.ToChar((int)('A' + modulo)) + columnLetter;
                idx = (idx - modulo) / 26;
            }
            return columnLetter;
        }
    }
}
