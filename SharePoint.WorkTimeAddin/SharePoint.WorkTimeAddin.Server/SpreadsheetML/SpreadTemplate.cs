using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Reflection;
using System.IO.Packaging;
using System.Data;
using DocumentFormat.OpenXml;

namespace SharePoint.WorkTimeAddin.SpreadsheetML
{
    /// <summary>
    /// スプレッドシート上の表示部品機能を提供します。
    /// </summary>
    /// <remarks>
    /// ■テンプレートの準備
    /// ① Excel上にテンプレートをデザインする
    /// <para/>
    /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
    /// さらにセルのスタイルにはバインド後に適用するものを指定します。
    /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
    /// <para/>
    /// ② 作成したテンプレートの範囲を名前付範囲する
    /// </remarks>
    public class SpreadTemplate
    {
        /// <summary>
        /// SpreadTemplateのCel情報を管理
        /// </summary>
        internal class Item
        {
            public string Text { get; set; }
            public string Formula { get; set; }
            public CellValues? DataType { get; set; }
            public uint? StyleIndex { get; set; }
        }

        /// <summary>
        /// テンプレートの左上のアドレス
        /// </summary>
        public SpreadsheetAddress BaseAddress { get; set; }

        internal Item[,] Cells { get; set; }
        internal List<SpreadsheetRange> MergeCells { get; set; }
        internal List<double?> Height { get; set; }

        /// <summary>
        /// テンプレートの列サイズ
        /// </summary>
        public int ColumnSize { get { return Cells.GetLength(0); } }

        /// <summary>
        /// テンプレートの行サイズ
        /// </summary>
        public int RowSize { get { return Cells.GetLength(1); } }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SpreadTemplate()
        {
            Height = new List<double?>();
        }

    }
}
