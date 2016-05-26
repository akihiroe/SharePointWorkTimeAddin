using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePoint.WorkTimeAddin.SpreadsheetML
{
    /// <summary>
    /// スプレッドシートの範囲を表現します
    /// </summary>
    public class SpreadsheetRange
    {
        /// <summary>
        /// 範囲の開始アドレス（左上）
        /// </summary>
        public SpreadsheetAddress Start { get; private set; }
        /// <summary>
        /// 範囲の終了アドレス（右下）
        /// </summary>
        public SpreadsheetAddress End { get; private set; }

        /// <summary>
        /// 範囲参照文字列（A1:B10、$A$10:$B:20、...）
        /// </summary>
        public string Range
        {
            get { return Start.Address + ":" + End.Address ?? Start.Address; }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="startColumnIndex">開始例番号（1～）</param>
        /// <param name="startRowIndex">開始行番号（1～）</param>
        /// <param name="endColumnIndex">終了例番号（1～）</param>
        /// <param name="endRowIndex">終了行番号（1～）</param>
        public SpreadsheetRange(int startColumnIndex, int startRowIndex, int endColumnIndex, int endRowIndex) 
            : this(new SpreadsheetAddress(startColumnIndex, startRowIndex), new SpreadsheetAddress(endColumnIndex, endRowIndex))
        {
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="start">開始アドレス</param>
        /// <param name="end">終了アドレス</param>
        public SpreadsheetRange(SpreadsheetAddress start, SpreadsheetAddress end)
        {
            if (start.ColumnIndex > end.ColumnIndex) throw new ArgumentOutOfRangeException("終了列番号は開始列番号以上を指定してください");
            if (start.RowIndex > end.RowIndex) throw new ArgumentOutOfRangeException("終了行番号は開始行番号以上を指定してください");
            this.Start = start;
            this.End = end;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="range">範囲参照文字列（A1:B10、$A$10:$B:20、...）</param>
        public SpreadsheetRange(string range)
        {
            var address2 = range.Split(':');
            if (address2.Length > 2 || address2.Length < 1) throw new ArgumentException("Rangeの形式が不正です。A1:B1の形式で入力してください。シート名は利用できません");
            this.Start = new SpreadsheetAddress(address2[0]);
            if (address2.Length == 2) this.End = new SpreadsheetAddress(address2[1]);
        }

        /// <summary>
        /// 範囲の移動変換を行います
        /// </summary>
        /// <param name="columnDelta">列移動量</param>
        /// <param name="rowDelta">行移動量</param>
        /// <returns>移動したスプレッドシート範囲</returns>
        /// <remarks>
        /// 移動は相対位置指定された行や列のみが対象になります。
        /// $で絶対指定された列や行は移動されません。
        /// </remarks>
        public SpreadsheetRange Translate(int columnDelta, int rowDelta)
        {
            return new SpreadsheetRange(Start.Translate(columnDelta, rowDelta), End.Translate(columnDelta, rowDelta));
        }

    }
}
