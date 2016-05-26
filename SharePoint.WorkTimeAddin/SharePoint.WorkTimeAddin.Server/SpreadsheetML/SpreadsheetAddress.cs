using System;
using System.Linq;
using System.Text.RegularExpressions;


namespace SharePoint.WorkTimeAddin.SpreadsheetML
{
    /// <summary>
    /// スプレッドシートのアドレスを表現します
    /// </summary>
    public class SpreadsheetAddress
    {
        /// <summary>
        /// 列番号（1～）
        /// </summary>
        public int ColumnIndex { get; private set; }

        /// <summary>
        /// 行番号（1～）
        /// </summary>
        public int RowIndex { get; private set; }

        /// <summary>
        /// 列位置が絶対指定($)か
        /// </summary>
        public bool ColumnAbsolute { get; private set; }

        /// <summary>
        /// 行位置が絶対指定($)か
        /// </summary>
        public bool RowAbsolute { get; private set; }

        /// <summary>
        /// 列文字列（A,B,C,...）
        /// </summary>
        public string ColumnLetter
        {
            get
            {
                if (_columnLetter != null) return _columnLetter;
                _columnLetter = SpreadsheetUtil.ConvertColumnLetter(this.ColumnIndex);
                return _columnLetter;

            }
        }
        private string _columnLetter;   //タイトループでの利用が予測されるためキャッシュ

        /// <summary>
        /// アドレス参照文字列（列文字列＋行番号形式の文字列。A1、C$10、$W$10）
        /// </summary>
        public string Address
        {
            get { return (this.ColumnAbsolute ? "$" : "") + ColumnLetter + (RowAbsolute ? "$" : "") + this.RowIndex.ToString(); }
        }


        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="columnAbsolute">列位置が絶対指定($)か</param>
        /// <param name="rowAbsolute">行位置が絶対指定($)か</param>
        public SpreadsheetAddress(int columnIndex, int rowIndex, bool columnAbsolute = false, bool rowAbsolute = false)
        {
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            ColumnAbsolute = columnAbsolute;
            RowAbsolute = rowAbsolute;
        }

        internal static Regex RengeRefExp = new Regex(@"^(?<ColAbs>\$)?(?<Col>[A-Z]+)(?<RowAbs>\$)?(?<Row>[0-9]+)$", RegexOptions.Compiled);

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="address">アドレス参照文字列（A1、C$10、$W$10）</param>
        public SpreadsheetAddress(string address)
        {
            if (address == null) throw new ArgumentNullException("address");
            var addressMatch = RengeRefExp.Match(address);
            if (!addressMatch.Success) throw new ArgumentOutOfRangeException(string.Format("{0}は不正なアドレスです", address), "adddress");
            ColumnAbsolute = "$" == addressMatch.Groups["ColAbs"].Value;
            ColumnIndex = SpreadsheetUtil.ConvertColumnIndex(addressMatch.Groups["Col"].Value);
            RowAbsolute = "$" == addressMatch.Groups["RowAbs"].Value;
            RowIndex = int.Parse(addressMatch.Groups["Row"].Value);
        }

        /// <summary>
        /// アドレスの移動変換を行います
        /// </summary>
        /// <param name="columnDelta">列移動量</param>
        /// <param name="rowDelta">行移動量</param>
        /// <returns>移動したスプレッドシート・アドレス</returns>
        /// <remarks>
        /// 移動は相対位置指定された行や列のみが対象になります。
        /// $で絶対指定された列や行は移動されません。
        /// </remarks>
        public SpreadsheetAddress Translate(int columnDelta, int rowDelta)
        {
            var newColumn = this.ColumnIndex + (ColumnAbsolute ? 0 : columnDelta);
            var newRow = this.RowIndex + (RowAbsolute ? 0 : rowDelta);
            if (newColumn <= 0 || newRow <= 0) throw new ArgumentException("移動先のアドレスが不正です");
            return new SpreadsheetAddress(newColumn, newRow, ColumnAbsolute, RowAbsolute);
        }
    }

}
