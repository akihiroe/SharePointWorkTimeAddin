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
    /// スプレッドシートの操作を支援します。
    /// </summary>
    public class SpreadsheetHelper : IDisposable
    {
        /// <summary>
        /// スプレッドシートドキュメント
        /// </summary>
        public SpreadsheetDocument Document { get; set; }

        /// <summary>
        /// カレントシート(Sheet)
        /// </summary>
        public Sheet CurrentSheet
        {
            get { return _currentSheet; }
            set
            {
                _currentSheet = value;
                sharedFormula = null;   //キャッシュクリア
                if (value == null)
                {
                    CurrentWorksheetPart = null;
                    CurrentWorksheet = null;
                    CurrentSheetData = null;
                }
                else
                {
                    CurrentWorksheetPart = (WorksheetPart)Document.WorkbookPart.GetPartById(CurrentSheet.Id);
                    CurrentWorksheet = CurrentWorksheetPart.Worksheet;
                    CurrentSheetData = CurrentWorksheet.GetFirstChild<SheetData>();
                }
            }
        }
        private Sheet _currentSheet;

        /// <summary>
        /// カレントシートのパーツ（WorksheetPart）
        /// </summary>
        public WorksheetPart CurrentWorksheetPart { get; private set; }

        /// <summary>
        /// カレントシートのシートデータ（SheetData）
        /// </summary>
        public SheetData CurrentSheetData { get; private set; }

        /// <summary>
        ///  カレントワークシート(Worksheet)
        /// </summary>
        public Worksheet CurrentWorksheet { get; private set; }


        /// <summary>
        /// ワークブックスタイル（WorkbookStylesPart）
        /// </summary>
        public WorkbookStylesPart StylePart { get; private set; }

        /// <summary>
        /// カレントシートのページ設定
        /// </summary>
        public PageSetup PageSetup
        {
            get
            {
                var setup = CurrentWorksheet.GetFirstChild<PageSetup>();
                if (setup == null)
                {
                    setup = new PageSetup();
                    CurrentWorksheet.InsertAfter(setup, CurrentWorksheet.GetFirstChild<PageMargins>());
                }
                return setup;
            }
        }

        /// <summary>
        /// カレントシートのページマージン設定
        /// </summary>
        public PageMargins PageMargins
        {
            get
            {
                return CurrentWorksheet.GetFirstChild<PageMargins>();//必ず存在する
            }
        }

        private MemoryStream documentStream;
        private Dictionary<uint, Tuple<string,string>> sharedFormula;

        #region WorkBook関連

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>
        /// 自動的に空シート(Sheet1)をもつワークブックが作成されます。
        /// </remarks>
        public SpreadsheetHelper()
        {
            using (var stream = this.GetType().Assembly.GetManifestResourceStream(this.GetType().Namespace + "." + "BlankTemplate.xlsx"))
            {
                if (stream == null) throw new ArgumentException(string.Format("リソース{0}.BlankTemplate.xlsxがありません", this.GetType().Namespace));
                var buffer = new byte[(int)stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                documentStream = new MemoryStream();
                documentStream.Write(buffer, 0, buffer.Length);
                Document = SpreadsheetDocument.Open(documentStream, true);
                StylePart = Document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                CurrentSheet = Document.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            }
        }

        /// <summary>
        /// 指定したファイル名でワークブックを読み込みます。
        /// </summary>
        /// <param name="filename">ファイル名</param>
        public SpreadsheetHelper(string filename)
        {
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                if (stream == null) throw new ArgumentException(string.Format("ファイル{0}がありません", filename));
                var buffer = new byte[(int)stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                documentStream = new MemoryStream();
                documentStream.Write(buffer, 0, buffer.Length);
                Document = SpreadsheetDocument.Open(documentStream, true);
                StylePart = Document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                CurrentSheet = Document.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            }
        }

        /// <summary>
        /// 指定したファイル名で保存します。
        /// </summary>
        /// <param name="filename">ファイル名</param>
        public void Save(string filename, bool updating = true)
        {
            //http://blogs.msdn.com/b/vsod/archive/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0.aspx
            CalculationChainPart calChainPart = this.Document.WorkbookPart.CalculationChainPart;
            if (calChainPart != null) this.Document.WorkbookPart.DeletePart(calChainPart);

            Document.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = updating;
            Document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = updating;


            Document.WorkbookPart.WorksheetParts.ToList().ForEach(x => x.Worksheet.Save());
            Document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().ToList().ForEach(x => x.Stylesheet.Save());
            Document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().ToList().ForEach(x => x.SharedStringTable.Save());
            Document.WorkbookPart.Workbook.Save();
            Document.Close();
            using (var file = new FileStream(filename, FileMode.Create))
            {
                documentStream.WriteTo(file);
                file.Close();
            }
        }

        #endregion

        #region Worksheet関連

        /// <summary>
        /// ワークシートを追加します。
        /// </summary>
        /// <param name="sheetname">シート名</param>
        /// <returns>追加した場合：true、既存のシートが存在する場合：false</returns>
        public bool AddWorksheet(string sheetname = null)
        {
            var sheet = Document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetname).FirstOrDefault();
            if (sheet == null)
            {
                uint newId = 1;
                if (Document.WorkbookPart.Workbook.Sheets.Count() > 0)
                {
                    newId = Document.WorkbookPart.Workbook.Sheets.Elements<Sheet>().Max(x => x.SheetId.Value) + 1;
                }
                string rId = "relId" + newId;
                if (string.IsNullOrEmpty(sheetname)) sheetname = string.Format("Sheet{0}", newId);
                WorksheetPart worksheet = Document.WorkbookPart.AddNewPart<WorksheetPart>(rId);
                var newSheet = Document.WorkbookPart.Workbook.Sheets.AppendChild<Sheet>(new Sheet() { Id = rId, SheetId = newId, Name = sheetname });
                worksheet.Worksheet = new Worksheet();
                worksheet.Worksheet.AppendChild<SheetData>(new SheetData());
                worksheet.Worksheet.AppendChild<PageMargins>(new PageMargins() { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 }); //標準設定
                CurrentSheet = newSheet;
                return true;
            }
            return false;
        }

        /// <summary>
        /// カレントのワークシートを削除します。
        /// </summary>
        /// <returns>削除した場合：true、シートが存在しない場合：false</returns>
        public bool RemoveWorksheet()
        {
            if (CurrentWorksheetPart == null) return false;
            Document.WorkbookPart.DeletePart(CurrentSheet.Id);
            Document.WorkbookPart.Workbook.Sheets.RemoveChild<Sheet>(CurrentSheet);
            CurrentSheet = null;
            return true;
        }

        /// <summary>
        /// カレントのワークシートを移動します。
        /// </summary>
        /// <param name="sheetName">シート名</param>
        /// <returns>移動した場合：true、シートが存在しない場合：false</returns>
        public bool MoveWorksheet(string sheetName)
        {
            var sheet = Document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (sheet == null) return false;
            CurrentSheet = sheet;
            return true;
        }

        /// <summary>
        /// 印刷領域を取得します
        /// </summary>
        /// <param name="printRange">印刷領域</param>
        public SpreadsheetRange GetPrintArea()
        {
            const string printAreaName = "_xlnm.Print_Area";
            var defines = Document.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (defines == null) return null;
            var sheetIdx = Document.WorkbookPart.Workbook.Descendants<Sheet>().TakeWhile(s => s.Name != CurrentSheet.Name).Count();
            var printArea = defines.Elements<DefinedName>().Where(x => x.Name == printAreaName && ConvertFromValue(x.LocalSheetId) == sheetIdx).FirstOrDefault();
            if (printArea != null)
            {
                var work = printArea.Text.Split('!');
                return new SpreadsheetRange(work.Length == 1 ? work[0] : work[1]);
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// 印刷領域を設定します
        /// </summary>
        /// <param name="printRange">印刷領域</param>
        public void SetPrintArea(SpreadsheetRange printRange)
        {
            const string printAreaName = "_xlnm.Print_Area";
            printRange = new SpreadsheetRange(
                new SpreadsheetAddress(printRange.Start.ColumnIndex, printRange.Start.RowIndex, true, true),
                new SpreadsheetAddress(printRange.End.ColumnIndex, printRange.End.RowIndex, true, true));
            var printRangeName = CurrentSheet.Name + "!" + printRange.Range;

            var defines = Document.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (defines == null)
            {
                defines = new DefinedNames();
                Document.WorkbookPart.Workbook.InsertAfter(defines, Document.WorkbookPart.Workbook.Descendants<Sheets>().First());
            }
            var sheetIdx = Document.WorkbookPart.Workbook.Descendants<Sheet>().TakeWhile(s => s.Name != CurrentSheet.Name).Count();
            var printArea = defines.Elements<DefinedName>().Where(x => x.Name == printAreaName && ConvertFromValue(x.LocalSheetId) == sheetIdx).FirstOrDefault();
            if (printArea != null)
            {
                printArea.Text = printRangeName;
            }
            else
            {
                defines.AppendChild(new DefinedName(printRangeName) { Name = printAreaName, LocalSheetId = (uint)sheetIdx });
            }
        }

        /// <summary>
        /// シートの列情報をコピーします。
        /// </summary>
        /// <param name="sheetName">コピー元のシート名</param>
        /// <returns>コピーした場合：true、シートが存在しない場合：false</returns>
        public bool CopySheetColumns(string sheetName)
        {
            var sheet = Document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (sheet == null) return false;
            var part = (WorksheetPart)Document.WorkbookPart.GetPartById(sheet.Id);
            CurrentWorksheet.InsertBefore<Columns>(part.Worksheet.GetFirstChild<Columns>().CloneElement<Columns>(), CurrentSheetData);
            return true;
        }

        /// <summary>
        /// シートの行を全て削除します
        /// </summary>
        /// <remarks>
        /// セルのマージ情報を全て削除されます。
        /// </remarks>
        public void ClearSheetRows()
        {
            CurrentSheetData.RemoveAllChildren<Row>();
            CurrentWorksheet.RemoveAllChildren<MergeCells>();
        }

        #endregion

        #region Cell関連

        /// <summary>
        /// 指定した行を取得します。
        /// </summary>
        /// <param name="index">行番号（1～）</param>
        /// <returns>行（Row）</returns>
        public Row GetRow(int index) //チューニング対象
        {
            if (index <= 0) throw new ArgumentOutOfRangeException("index");

            //行追加処理を早く実行するための仕組みを追加
            var last = CurrentSheetData.LastChild as Row;
            if (last != null)
            {
                if (last.RowIndex == index) return last;
                if (last.RowIndex < index)
                {
                    var lastNewRow = new Row() { RowIndex = (uint)index };
                    CurrentSheetData.AppendChild<Row>(lastNewRow);
                    return lastNewRow;
                }
            }

            int pos = 0;
            foreach (var row in CurrentSheetData.Elements<Row>())
            {
                if (row.RowIndex.Value == index) return row;
                if (row.RowIndex.Value > index) break;
                pos++;
            }
            var newRow = new Row() { RowIndex = (uint)index };
            CurrentSheetData.InsertAt(newRow, pos);
            return newRow;

            //LINQ Simple版
            //var row = CurrentSheetData.Elements<Row>().Where(r => r.RowIndex.Value == index).FirstOrDefault();
            //if (row == null)
            //{
            //    row = new Row() { RowIndex = (uint)index };
            //    CurrentSheetData.InsertAt(row, CurrentSheetData.Elements<Row>().Count(x => x.RowIndex.Value < index));
            //}
            //return row;
        }


        /// <summary>
        /// 指定した列を取得します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <returns>列（Column）</returns>
        public Column GetColumn(int columnIndex)
        {
            var columns = CurrentWorksheet.GetFirstChild<Columns>();
            if (columns == null) columns = CurrentWorksheet.InsertBefore<Columns>(new Columns(), CurrentSheetData);
            var column = columns.Elements<Column>().Where(x => columnIndex >= x.Min && columnIndex <= x.Max).FirstOrDefault();
            if (column != null)
            {
                if (column.Min != columnIndex)
                {
                    var newColumn = column.CloneElement<Column>();
                    columns.InsertBefore<Column>(newColumn, column);
                    newColumn.Max = (uint)columnIndex - 1;
                    column.Min = (uint)columnIndex;
                }
                if (column.Max != columnIndex)
                {
                    var newColumn = column.CloneElement<Column>();
                    columns.InsertAfter<Column>(newColumn, column);
                    newColumn.Min = (uint)columnIndex + 1;
                    column.Max = (uint)columnIndex;
                }
                return column;
            }
            else
            {
                column = columns.Elements<Column>().Where(x => x.Max < columnIndex).FirstOrDefault();
                var newColumn = column == null ? new Column() : column.CloneElement<Column>();
                newColumn.Max = (uint)columnIndex;
                newColumn.Min = (uint)columnIndex;
                columns.InsertBefore<Column>(newColumn, column);
                return newColumn;
            }
        }

        /// <summary>
        /// 指定したセルを取得します。
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <returns>セル（Cell）</returns>
        public Cell GetCell(string address)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return GetCell(ssAdrs.ColumnIndex, ssAdrs.RowIndex);
        }

        
        /// <summary>
        /// 指定したセルを取得します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <returns>セル（Cell）</returns>
        public Cell GetCell(int columnIndex, int rowIndex) //チューニング対象
        {
            var row = GetRow(rowIndex);
            int rowStringLength = rowIndex.ToString().Length;


            //セル追加処理を早く実行するための仕組みを追加
            var last = row.LastChild as Cell;
            if (last != null)
            {
                string cellText = last.CellReference.Value;
                string columnLetter = cellText.Substring(0, cellText.Length - rowStringLength);
                int idx = SpreadsheetUtil.ConvertColumnIndex(columnLetter);
                if (idx == columnIndex) return last;
                if (idx < columnIndex)
                {
                    var lastNewCell = new Cell() { CellReference = SpreadsheetUtil.ConvertColumnLetter(columnIndex) + rowIndex.ToString() };
                    row.AppendChild<Cell>(lastNewCell);
                    return lastNewCell;
                }
            }

            int pos = 0;
            foreach (Cell cell in row.Elements<Cell>())
            {
                string cellText = cell.CellReference.Value;
                string columnLetter = cellText.Substring(0, cellText.Length - rowStringLength);
                int idx = SpreadsheetUtil.ConvertColumnIndex(columnLetter);
                if (idx == columnIndex) return cell;
                if (idx > columnIndex) break;
                pos++;
            }
            var newCell = new Cell() { CellReference = SpreadsheetUtil.ConvertColumnLetter(columnIndex) + rowIndex.ToString() };
            row.InsertAt(newCell, pos);
            return newCell;

            //LINQ Simple版
            //var row = GetRow(rowIndex);
            //var address = new SpreadAddress(colIndex, rowIndex);
            //var cell = row.Elements<Cell>().Where(x => x.CellReference.Value == address.Address).FirstOrDefault();
            //if (cell == null)
            //{
            //    cell = new Cell() { CellReference = address.Address };
            //    row.InsertAt(cell, row.Elements<Cell>().Count(x => new SpreadAddress(x.CellReference).ColumnIndex < address.ColumnIndex));
            //}
            //return cell;
        }

        /// <summary>
        /// 指定したセルの値を取得します。
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <returns>セル（Cell）</returns>
        /// <remarks>
        /// データ型の判断はセルのデータ型とスタイルから判断します。
        /// スプレッドシートでは日付型は数値として保持さスタイルが日付型として指定することで日付表示されています。
        /// </remarks>
        public object GetCellValue(string address, CellValues? dataType = null)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return GetCellValue(ssAdrs.ColumnIndex, ssAdrs.RowIndex, dataType);
        }

        /// <summary>
        /// 指定したセルの値を取得します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="dataType">取得するデータの種類。nullを指定する自動判断</param>
        /// <returns>セル（Cell）</returns>
        /// <remarks>
        /// データ型の判断はセルのデータ型とスタイルから判断します。
        /// スプレッドシートでは日付型は数値として保持さスタイルが日付型として指定することで日付表示されています。
        /// </remarks>
        public object GetCellValue(int columnIndex, int rowIndex, CellValues? dataType = null)
        {
            var cell = GetCell(columnIndex, rowIndex);
            if (cell == null) return null;
            if (dataType == null) dataType = ConvertFromEnumValue(cell.DataType);
            return ConvertCellValue(cell.CellValue, dataType, GetCellFormat(cell));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <param name="value">設定値</param>
        /// <param name="styleIndex">スタイル番号</param>
        /// <returns>設定したセル</returns>
        public Cell SetCellValue(string address, object value, uint? styleIndex = null)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return SetCellValue(ssAdrs.ColumnIndex, ssAdrs.RowIndex,value, styleIndex);
        }
        /// <summary>
        /// 指定したセルに値を設定します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="value">設定値</param>
        /// <param name="styleIndex">スタイル番号</param>
        /// <returns>設定したセル</returns>
        public Cell SetCellValue(int columnIndex, int rowIndex, object value, uint? styleIndex = null)
        {
            return SetCellValue(GetCell(columnIndex, rowIndex), value, styleIndex);
        }

        /// <summary>
        /// 指定したセルを更新します
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <returns>設定したセル</returns>
        public Cell UpdateCellValue(Cell cell)
        {
             cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            return cell;
        }
        /// <summary>
        /// 指定したセルに値を設定します。
        /// </summary>
        /// <param name="cell">値を設定するセル</param>
        /// <param name="value">設定値</param>
        /// <param name="styleIndex">スタイル番号</param>
        /// <param name="overrideDateType">DataTypeを上書きするかどうか</param>
        /// <returns></returns>
        internal Cell SetCellValue(Cell cell, object value, uint? styleIndex = null, bool overrideDateType = true, Type dataType=null)
        {
            if (value == null)
            {
                cell.CellValue = null;
            }
            else
            {
                cell.CellValue = ConvertCellValue(value, value.GetType());
                cell.DataType = GetCellValues(dataType ?? value.GetType());
                if (styleIndex != null)
                {
                    if (styleIndex != null) cell.StyleIndex = styleIndex;
                }
                else
                {
                    if (overrideDateType && value.GetType() == typeof(DateTime) && IsNullOrNoValue(cell.StyleIndex)) cell.StyleIndex = GetBuildinStyleIndex(14);
                }
            }
            return cell;
        }

        /// <summary>
        /// 指定したセルに式を設定します。
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <param name="formula">式（例 SUM(A1:B2) ）</param>
        /// <param name="styleIndex">スタイル番号</param>
        /// <returns></returns>
        public Cell SetCellFormula(string address, string formula, uint? styleIndex = null)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return SetCellFormula(ssAdrs.ColumnIndex, ssAdrs.RowIndex, formula, styleIndex);
        }
        /// <summary>
        /// 指定したセルに式を設定します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="formula">式（例 SUM(A1:B2) ）</param>
        /// <param name="styleIndex">スタイル番号</param>
        /// <returns></returns>
        public Cell SetCellFormula(int columnIndex, int rowIndex, string formula, uint? styleIndex = null)
        {
            var cell = GetCell(columnIndex, rowIndex);
            if (formula == null)
            {
                cell.Parent.RemoveChild(cell);
            }
            else
            {
                cell.CellValue = null;
                cell.CellFormula = new CellFormula(formula);
                if (styleIndex != null) cell.StyleIndex = styleIndex;
            }
            return cell;
        }

        ////性能の観点でstyleIndexを意識させたほうが良いようなので以下のAPIは削除予定
        //public Cell SetCellValue(int colIndex, int rowIndex, object value, SpreadStyle style)
        //{
        //    return SetCellValue(colIndex, rowIndex, value, style == null ? (uint?)null : GetStyleIndex(style));
        //}
        //public Cell SetCellFormula(int colIndex, int rowIndex, string formula, SpreadStyle style)
        //{
        //    return SetCellFormula(colIndex, rowIndex, formula, style == null ? (uint?)null : GetStyleIndex(style));
        //}

        /// <summary>
        /// セルをマージします。
        /// </summary>
        /// <param name="startColumnIndex">開始列番号（1～）</</param>
        /// <param name="startRowIndex">開始行番号（1～）</param>
        /// <param name="endColumnIndex">終了列番号（1～）</param>
        /// <param name="endRowIndex">終了行番号（1～）</param>
        public void MergeCells(int startColumnIndex, int startRowIndex, int endColumnIndex, int endRowIndex)
        {
            MergeCells(new SpreadsheetRange(new SpreadsheetAddress(startColumnIndex, startRowIndex), new SpreadsheetAddress(endColumnIndex, endRowIndex)));
        }

        /// <summary>
        /// セルをマージします。
        /// </summary>
        /// <param name="range">マージ範囲</param>
        public void MergeCells(SpreadsheetRange range)
        {
            var worksheet = CurrentWorksheet;
            var mergeCells = worksheet.GetFirstChild<MergeCells>();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells() { Count = 0 };
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            }
            var cell = mergeCells.Elements<MergeCell>().Where(x => x.Reference.Value.Equals(range.Range, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (cell != null) return;
            var newMergeCell = new MergeCell() { Reference = range.Range };
            mergeCells.AppendChild<MergeCell>(newMergeCell);
            mergeCells.Count++;
        }

        /// <summary>
        /// セルマージで非表示のセル
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public bool IsHideCellByMergeCells(int columnIndex, int rowIndex)
        {
            var worksheet = CurrentWorksheet;
            var mergeCells = worksheet.GetFirstChild<MergeCells>();
            if (mergeCells == null) return false;
            foreach(var cell in mergeCells.Elements<MergeCell>())
            {
                var range = new SpreadsheetRange(cell.Reference.Value);
                if (range.Start.ColumnIndex > columnIndex) continue;
                if (range.End.ColumnIndex < columnIndex) continue;
                if (range.Start.RowIndex > rowIndex) continue;
                if (range.End.RowIndex < rowIndex) continue;
                if (range.Start.ColumnIndex != columnIndex
                    || range.Start.RowIndex != rowIndex) return true;
            }
            return false;
        }

        /// <summary>
        /// セルのマージを解除します。
        /// </summary>
        /// <param name="range">解除するマージ範囲</param>
        /// <returns>解除した場合：true、解除するマージされたセルが存在しない場合：false</returns>
        public bool UnmergeCells(SpreadsheetRange range)
        {
            var worksheet = CurrentWorksheet;
            var cells = worksheet.GetFirstChild<MergeCells>();
            if (cells == null) return false;
            var cell = cells.Elements<MergeCell>().Where(x => x.Reference.Value.Equals(range.Range, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (cell == null) return false;
            cells.RemoveChild(cell);
            cells.Count--;
            if (cells.Count == 0) CurrentWorksheet.RemoveChild<MergeCells>(cells);
            return true;
        }

        private CellFormat GetCellFormat(Cell cell)
        {
            return cell.StyleIndex != null ? (CellFormat)StylePart.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value] : null;
        }

        private CellValue ConvertCellValue(object value, Type valueType)
        {
            if (valueType == typeof(DateTime)) return new CellValue(((DateTime)value).ToOADate().ToString());
            if (valueType == typeof(bool)) return new CellValue(((bool)value) ? "1" : "0");
            return new CellValue(value.ToString());
        }

        private object ConvertCellValue(CellValue value, CellValues? valueType, CellFormat format)
        {
            if (value == null) return null;
            //日付フォーマット http://www.documentinteropinitiative.org/implnotes/ECMA-376/5788b1ab-cb6d-4973-8c69-0d5446d3b36e.aspx
            if (!valueType.HasValue && format != null && (
                (format.NumberFormatId >= 14 && format.NumberFormatId <= 22) ||
                (format.NumberFormatId >= 27 && format.NumberFormatId <= 36) ||
                (format.NumberFormatId >= 50 && format.NumberFormatId <= 58))) return DateTime.FromOADate(double.Parse(value.Text));

            if (valueType == CellValues.Date) return DateTime.FromOADate(double.Parse(value.Text));
            if (valueType == CellValues.SharedString) return GetSharedStringValue(value);
            if (valueType == CellValues.Boolean) return value.Text == "1" ? true : value.Text == "0" ? false : Boolean.Parse(value.Text);
            if (valueType == CellValues.Number) return Decimal.Parse(value.Text);
            return value.Text;
        }

        private string GetSharedStringValue(CellValue value)
        {
            var item = (SharedStringItem)Document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[Int32.Parse(value.Text)];
            return item.Text != null ? item.Text.Text : null;
        }

        private CellValues GetCellValues(Type property)
        {
            if (property == typeof(String)) return CellValues.String;
            if (property == typeof(int)) return CellValues.Number;
            if (property == typeof(long)) return CellValues.Number;
            if (property == typeof(decimal)) return CellValues.Number;
            if (property == typeof(double)) return CellValues.Number;
            if (property == typeof(float)) return CellValues.Number;
            if (property == typeof(byte)) return CellValues.Number;
            if (property == typeof(uint)) return CellValues.Number;
            if (property == typeof(ulong)) return CellValues.Number;
            if (property == typeof(Boolean)) return CellValues.Boolean;
            if (property == typeof(DateTime)) return CellValues.Number;
            if (property == typeof(TimeSpan)) return CellValues.Number;
            return CellValues.String;
        }

        #endregion

        #region Style関連

        /// <summary>
        /// 既定のスタイルを取得します。
        /// </summary>
        /// <returns>スタイル</returns>
        public SpreadsheetStyle GetDefaultStyle()
        {
            return GetStyle((CellFormat)StylePart.Stylesheet.CellStyleFormats.FirstChild);
        }

        /// <summary>
        /// 指定したセルのスタイルを取得します。
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <returns>スタイル</returns>
        public SpreadsheetStyle GetCellStyle(string address)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return GetCellStyle(ssAdrs.ColumnIndex, ssAdrs.RowIndex);
        }

        /// <summary>
        /// 指定したセルのスタイルを取得します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <returns>スタイル</returns>
        public SpreadsheetStyle GetCellStyle(int columnIndex, int rowIndex)
        {
            return GetStyle(GetCellFormat(GetCell(columnIndex, rowIndex)));
        }

        /// <summary>
        /// 指定したセルのスタイル番号を取得します。
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <returns>スタイル番号。セル上にスタイル番号が設定されていない場合はnull</returns>
        public uint? GetCellStyleIndex(int columnIndex, int rowIndex)
        {
            return ConvertFromValue(GetCell(columnIndex, rowIndex).StyleIndex);
        }

        /// <summary>
        /// 指定したセルにスタイルを設定します
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <param name="style">スタイル</param>
        public void SetCellStyle(string address, SpreadsheetStyle style)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            SetCellStyle(ssAdrs.ColumnIndex, ssAdrs.RowIndex, style);
        }

        /// <summary>
        /// 指定したセルにスタイルを設定します
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="style">スタイル</param>
        public void SetCellStyle(int columnIndex, int rowIndex, SpreadsheetStyle style)
        {
            GetCell(columnIndex, rowIndex).StyleIndex = LookupStyleIndex(style);
        }

        /// <summary>
        /// 指定したセルにスタイル番号を設定します
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <param name="styleIndex">スタイル番号</param>
        public void SetCellStyleIndex(int columnIndex, int rowIndex, uint? styleIndex)
        {
            if (styleIndex.HasValue) GetCell(columnIndex, rowIndex).StyleIndex = styleIndex.Value;
        }

        /// <summary>
        /// 指定したスタイルのスタイル番号を検索します。
        /// </summary>
        /// <param name="style">スタイル</param>
        /// <returns>スタイル番号</returns>
        /// <remarks>
        /// Fontやボーダーなどのスタイルプロパティの設定値によってスタイル番号は変化します。
        /// したがって、同じSpreadStyleオブジェクトを指定しても同じスタイル番号が返るとは限りません。
        /// </remarks>
        public uint LookupStyleIndex(SpreadsheetStyle style)
        {
            uint fontIndex = GetFontIndex(style.fontElement);
            uint fillIndex = GetFillIndex(style.fillElement);
            uint borderIndex = GetBorderIndex(style.borderElement);
            uint numberFormatIndex = GetNumberFormatIndex(style.numberingFormatElement);

            uint sidx = 0;
            foreach (CellFormat cellFormat in StylePart.Stylesheet.CellFormats)
            {
                if (CompareValue(cellFormat.FontId, fontIndex) && CompareValue(cellFormat.FillId, fillIndex)
                    && CompareValue(cellFormat.BorderId, borderIndex) && CompareValue(cellFormat.NumberFormatId, numberFormatIndex))
                {

                    if (new SpreadElement<Alignment>(cellFormat.Alignment).Equals(style.alignmentElement)) return sidx;
                }
                sidx++;
            }
            CellFormat newFormat = new CellFormat() { NumberFormatId = numberFormatIndex, FontId = fontIndex, FillId = fillIndex, BorderId = borderIndex };
            if (style.alignmentElement != null) newFormat.Alignment = style.alignmentElement.CloneElement();

            newFormat.ApplyFont = new BooleanValue(true);
            if (newFormat.NumberFormatId.Value > 0) newFormat.ApplyNumberFormat = true;
            if (newFormat.FillId.Value > 0) newFormat.ApplyFill = true;
            if (newFormat.BorderId.Value > 0) newFormat.ApplyBorder = true;
            if (newFormat.Alignment != null) newFormat.ApplyAlignment = true;

            StylePart.Stylesheet.CellFormats.AppendChild<CellFormat>(newFormat);
            StylePart.Stylesheet.CellFormats.Count++;
            return sidx;

        }


        private SpreadsheetStyle GetStyle(CellFormat cellStyleFormat)
        {
            var font = (Font)StylePart.Stylesheet.Fonts.ChildElements[(int)cellStyleFormat.FontId.Value];
            var fill = (Fill)StylePart.Stylesheet.Fills.ChildElements[(int)cellStyleFormat.FillId.Value];
            var border = (Border)StylePart.Stylesheet.Borders.ChildElements[(int)cellStyleFormat.BorderId.Value];

            var style = new SpreadsheetStyle
            {
                fontElement = new SpreadElement<Font>(font),
                fillElement = new SpreadElement<Fill>(fill),
                borderElement = new SpreadElement<Border>(border),
                numberingFormatElement = new SpreadElement<NumberingFormat> { Element = new NumberingFormat { NumberFormatId = cellStyleFormat.NumberFormatId } },
            };
            return style;
        }


        internal uint GetBuildinStyleIndex(uint numberFormatIndex)
        {
            uint idx = 0;

            foreach (CellFormat cellFormat in StylePart.Stylesheet.CellFormats)
            {
                if (cellFormat.NumberFormatId.Value == numberFormatIndex) return idx;
                idx++;
            }
            CellFormat newFormat = new CellFormat() { NumberFormatId = numberFormatIndex, ApplyNumberFormat = true };
            StylePart.Stylesheet.CellFormats.AppendChild<CellFormat>(newFormat);
            StylePart.Stylesheet.CellFormats.Count++;
            return idx;
        }

        private uint GetFontIndex(SpreadElement<Font> styleFont)
        {
            uint idx = 0;
            foreach (Font font in StylePart.Stylesheet.Fonts)
            {
                if (new SpreadElement<Font>(font).Equals(styleFont)) return idx;
                idx++;
            }
            StylePart.Stylesheet.Fonts.AppendChild<Font>(styleFont.CloneElement());
            StylePart.Stylesheet.Fonts.Count++;
            return idx;

        }

        private uint GetBorderIndex(SpreadElement<Border> styleBorder)
        {
            uint idx = 0;
            foreach (Border border in StylePart.Stylesheet.Borders)
            {
                if (new SpreadElement<Border>(border).Equals(styleBorder)) return idx;
                idx++;
            }
            StylePart.Stylesheet.Borders.AppendChild<Border>(styleBorder.CloneElement());
            StylePart.Stylesheet.Borders.Count++;
            return idx;
        }

        private uint GetFillIndex(SpreadElement<Fill> styleFill)
        {
            uint idx = 0;
            foreach (Fill fill in StylePart.Stylesheet.Fills)
            {
                if (new SpreadElement<Fill>(fill).Equals(styleFill)) return idx;
                idx++;
            }
            StylePart.Stylesheet.Fills.AppendChild<Fill>(styleFill.CloneElement());
            StylePart.Stylesheet.Fills.Count++;
            return idx;
        }

        private uint GetNumberFormatIndex(SpreadElement<NumberingFormat> styleFormat)
        {
            if (styleFormat.Element.FormatCode == null || !styleFormat.Element.FormatCode.HasValue) return styleFormat.Element.NumberFormatId;
            if (StylePart.Stylesheet.NumberingFormats == null) StylePart.Stylesheet.NumberingFormats = new NumberingFormats() { Count = 0 };

            uint idx = 0;
            foreach (NumberingFormat format in StylePart.Stylesheet.NumberingFormats)
            {
                if (format.FormatCode == styleFormat.Element.FormatCode.Value) return idx;
                idx++;
            }
            idx += 164; //開始オフセット
            StylePart.Stylesheet.NumberingFormats.AppendChild<NumberingFormat>(
                new NumberingFormat() { FormatCode = styleFormat.Element.FormatCode, NumberFormatId = idx });
            StylePart.Stylesheet.NumberingFormats.Count++;
            return idx;
        }

        #endregion

        #region Formula関連


        /// <summary>
        /// 式（文字列）が参照しているセルを移動させた式（文字列）を取得します。
        /// </summary>
        /// <param name="formula">元の式（文字列）</param>
        /// <param name="columnDelta">列移動量</param>
        /// <param name="rowDelta">行移動量</param>
        /// <returns>移動された式（文字列）</returns>
        /// <remarks>
        /// セルの移動は相対位置指定された行や列のみが対象になります。
        /// $で絶対指定された列や行は移動されません。
        /// 例）行移動+5、列移動+1に移動する場合。SUM(A1:A20)→SUM(B6:B26)、SUM(A$1:A20)→SUM(B$1:B26)
        /// </remarks>
        public string TranslateFormula(string formula, int columnDelta, int rowDelta)
        {

            char[] token = { '(', ')', '+', '-', '*', '&', '/', ' ', '=', ',', '!', ':' };

            formula = formula + " "; //最後をTokenで終わらせる
            string transalateExpresssion = "";
            int pre = 0;
            for (int idx = 1; idx < formula.Length; idx++)
            {
                if (token.Contains(formula[idx]))
                {
                    string word = formula.Substring(pre, idx - pre);
                    transalateExpresssion += TransalateWordForAddress(word, columnDelta, rowDelta);
                    if (idx == formula.Length - 1) break;
                    transalateExpresssion += formula.Substring(idx, 1);
                    idx++;
                    pre = idx;
                }
            }
            return transalateExpresssion;
        }

        private string TransalateWordForAddress(string address, int colDelta, int rowDelta)
        {
            return SpreadsheetAddress.RengeRefExp.IsMatch(address) ? new SpreadsheetAddress(address).Translate(colDelta, rowDelta).Address : address;
        }

        /// <summary>
        /// 指定したセルの式を取得します
        /// </summary>
        /// <param name="address">セルアドレス（例 B1、F$1）</param>
        /// <returns>式</returns>
        public string GetFormula(string address)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return GetFormula(ssAdrs.ColumnIndex, ssAdrs.RowIndex);
        }

        /// <summary>
        /// 指定したセルの式を取得します
        /// </summary>
        /// <param name="columnIndex">列番号（1～）</param>
        /// <param name="rowIndex">行番号（1～）</param>
        /// <returns>式</returns>
        public string GetFormula(int columnIndex, int rowIndex)
        {
            return GetFormula(GetCell(columnIndex, rowIndex));
        }

        /// <summary>
        /// 指定したセルの式を取得します
        /// </summary>
        /// <param name="cell">セル</param>
        /// <returns>式</returns>
        public string GetFormula(Cell cell)
        {
            if (cell.CellFormula == null) return null;
            var formula = cell.CellFormula;
            if (!IsNullOrNoValue(formula.SharedIndex)) return GetSharedFormula(formula.SharedIndex.Value, cell.CellReference);
            return cell.CellFormula == null ? null : cell.CellFormula.Text;
        }

        private string GetSharedFormula(uint index, string target)
        {
            if (sharedFormula == null)
            {
                sharedFormula = new Dictionary<uint, Tuple<string, string>>();
                foreach (var cell in CurrentSheetData.Descendants<Cell>())
                {
                    if (cell.CellFormula != null && !IsNullOrNoValue(cell.CellFormula.SharedIndex) && cell.CellFormula.Reference != null)
                    {
                        sharedFormula.Add(cell.CellFormula.SharedIndex, new Tuple<string, string>(cell.CellFormula.Text, cell.CellReference));
                    }
                }
            }
            var baseAddress = new SpreadsheetAddress(sharedFormula[index].Item2);
            var targetAddress = new SpreadsheetAddress(target);
            return TranslateFormula(sharedFormula[index].Item1, targetAddress.ColumnIndex - baseAddress.ColumnIndex, targetAddress.RowIndex - baseAddress.RowIndex);
        }

        #endregion

        #region その他

        /// <summary>
        /// 名前付範囲を取得します。
        /// </summary>
        /// <param name="name">名前付範囲の名前</param>
        /// <returns>名前付範囲（DefinedName）</returns>
        public DefinedName GetDefinedName(string name)
        {
            var defines = Document.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            return defines != null ? defines.Elements<DefinedName>().FirstOrDefault(x => x.Name == name) : null;
        }

        public DefinedName SetDefineName(string name, string range)
        {
            var defines = Document.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (defines == null)
            {
                defines = new DefinedNames();
                Document.WorkbookPart.Workbook.InsertAfter(defines, Document.WorkbookPart.Workbook.Descendants<Sheets>().First());
            }

            var define = defines.Elements<DefinedName>().FirstOrDefault(x => x.Name == name);
            if (define != null)
            {
                define.Text = range;
            }
            else
            {
                defines.AppendChild(new DefinedName(range) { Name = name });
            }
            return define;
        }

        private T? ConvertFromEnumValue<T>(EnumValue<T> item) where T : struct
        {
            if (item == null) return null;
            if (!item.HasValue) return null;
            return item.Value;
        }
        private EnumValue<T> ConvertToEnumValue<T>(T? item) where T : struct
        {
            if (!item.HasValue) return null;
            return item.Value;
        }

        private T? ConvertFromSimpleValue<T>(OpenXmlSimpleValue<T> item) where T : struct
        {
            if (item == null) return null;
            if (!item.HasValue) return null;
            return (T)item.Value;
        }

        private uint? ConvertFromValue(UInt32Value value)
        {
            if (value == null) return null;
            return value.Value;
        }

        private bool CompareValue(UInt32Value value1, UInt32 value2)
        {
            if (value1 == null) return true;
            return value1.Value == value2;
        }

        private bool IsNullOrNoValue(UInt32Value value)
        {
            if (value == null) return true;
            return !value.HasValue;
        }

        #endregion

        #region テンプレート機能

        /// <summary>
        /// テンプレートを取得します。
        /// </summary>
        /// <param name="name">名前付範囲の名前</param>
        /// <returns>テンプレート</returns>
        public SpreadTemplate GetTemplate(string name)
        {
            var current = this.CurrentSheet;
            try
            {
                var rangeStr = GetDefinedName(name);
                if (rangeStr == null || rangeStr.Text == null) return null;
                var sheetRange = rangeStr.Text.Split('!');
                if (sheetRange.Length > 2) throw new ArgumentOutOfRangeException("name");

                if (sheetRange.Length == 2) MoveWorksheet(sheetRange[0]);
                var range = new SpreadsheetRange(sheetRange.Length == 1 ? sheetRange[0] : sheetRange[1]);
                var template = new SpreadTemplate();

                template.Cells = new SpreadTemplate.Item[range.End.ColumnIndex - range.Start.ColumnIndex + 1, range.End.RowIndex - range.Start.RowIndex + 1];
                for (int rowIdx = range.Start.RowIndex; rowIdx <= range.End.RowIndex; rowIdx++)
                {
                    for (int colIdx = range.Start.ColumnIndex; colIdx <= range.End.ColumnIndex; colIdx++)
                    {
                        var cell = GetCell(colIdx, rowIdx);
                        if (cell == null) continue;
                        var item = new SpreadTemplate.Item();
                        item.DataType = ConvertFromEnumValue(cell.DataType);
                        item.StyleIndex = ConvertFromValue(cell.StyleIndex);
                        item.Formula = GetFormula(cell);
                        if (cell.CellValue != null)
                        {
                            if (ConvertFromEnumValue(cell.DataType) == CellValues.SharedString)
                            {
                                item.Text = GetSharedStringValue(cell.CellValue);
                                item.DataType = CellValues.String;
                            }
                            else
                            {
                                item.Text = cell.CellValue.Text;
                            }
                        }
                        template.Cells[colIdx - range.Start.ColumnIndex, rowIdx - range.Start.RowIndex] = item;
                    }
                    var height = GetRow(rowIdx).Height;
                    template.Height.Add(ConvertFromSimpleValue(height));
                }
                template.MergeCells = FindMergeCells(range);
                template.BaseAddress = range.Start;
                return template;
            }
            finally
            {
                CurrentSheet = current;
            }
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <param name="address">配置する左上セルアドレス（例 B1、F$1）</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSizeを追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData(string address, object bindData, SpreadTemplate template, bool bindOnly = false)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return PutData(address, bindData, template, bindOnly);
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <typeparam name="T">リストアイテムの型</typeparam>
        /// <param name="address">配置する左上セルアドレス（例 B1、F$1）</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSize*行数を追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData<T>(string address, IEnumerable<T> bindData, SpreadTemplate template, bool bindOnly = false)
        {
            var ssAdrs = new SpreadsheetAddress(address);
            return PutData<T>(address, bindData, template, bindOnly);
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <param name="address">配置する左上</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSizeを追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData(SpreadsheetAddress address, object bindData, SpreadTemplate template, bool bindOnly = false)
        {
            return PutData(address.ColumnIndex, address.RowIndex, bindData, template, bindOnly);
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <typeparam name="T">リストアイテムの型</typeparam>
        /// <param name="address">配置する左上</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSize*行数を追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData<T>(SpreadsheetAddress address, IEnumerable<T> bindData, SpreadTemplate template, bool bindOnly = false)
        {
            return PutData<T>(address.ColumnIndex, address.RowIndex, bindData, template, bindOnly);
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <typeparam name="T">リストアイテムの型</typeparam>
        /// <param name="columnIndex">配置する左上の列番号</param>
        /// <param name="rowIndex">配置する左上の行番号</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSize*行数を追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData<T>(int columnIndex, int rowIndex, IEnumerable<T> bindData, SpreadTemplate template, bool bindOnly = false)
        {
            bindContext = null;
            if (bindData != null)
            {
                bindContext = typeof(T).GetProperties().Where(x => x.GetIndexParameters().Length == 0).ToDictionary(x => x.Name, x => x);
            }
            foreach (var bindItem in bindData)
            {
                PutDataInternl(columnIndex, rowIndex, bindItem, template, bindOnly);
                rowIndex += template.RowSize;
            }
            return rowIndex;
        }

        /// <summary>
        /// テンプレートを指定してデータを設定します。
        /// </summary>
        /// <param name="columnIndex">配置する左上の列番号</param>
        /// <param name="rowIndex">配置する左上の行番号</param>
        /// <param name="bindData">テンプレートにバインドするデータ</param>
        /// <param name="template">テンプレート</param>
        /// <param name="bindOnly">false：バインドした値だけ設定、true：テンプレートのスタイルも合わせて設定</param>
        /// <returns>設定を行った次行位置（addressのRowIndexにTemplateのRowSizeを追加した値）</returns>
        /// <remarks>
        /// テンプレート上のバインド対象のセルには[bindDataのプロパティ名]形式の文字列を記述します。例）[Now],[OrderNo]<para/>
        /// さらにセルのスタイルにはバインド後に適用するものを指定します。
        /// 特に日付を設定する場合、適切なスタイルが設定されていないと数値で表示されます。
        /// </remarks>
        public int PutData(int columnIndex, int rowIndex, object bindData, SpreadTemplate template, bool bindOnly = false)
        {
            bindContext = null;
            if (bindData != null)
            {
                bindContext = bindData.GetType().GetProperties().Where(x => x.GetIndexParameters().Length == 0).ToDictionary(x => x.Name, x => x);
            }
            PutDataInternl(columnIndex, rowIndex, bindData, template, bindOnly);
            return rowIndex + template.RowSize;
        }

        internal void PutDataInternl(int columnIndex, int rowIndex, object bindData, SpreadTemplate template, bool bindOnly)
        {
            int colDelta = columnIndex - template.BaseAddress.ColumnIndex;
            int rowDelta = rowIndex - template.BaseAddress.RowIndex;
            for (int rowIdx = 0; rowIdx < template.Cells.GetLength(1); rowIdx++)
            {
                for (int colIdx = 0; colIdx < template.Cells.GetLength(0); colIdx++)
                {
                    var item = template.Cells[colIdx, rowIdx];
                    var cell = GetCell(colIdx + columnIndex, rowIdx + rowIndex);

                    if (!bindOnly)
                    {
                        cell.DataType = ConvertToEnumValue(item.DataType);
                        if (item.Formula != null)
                        {
                            cell.CellFormula = new CellFormula(TranslateFormula(item.Formula, colDelta, rowDelta));
                            cell.CellValue = null;
                        }
                        else
                        {
                            BindValue(item, cell, bindData, bindOnly);
                        }
                        if (item.StyleIndex.HasValue) cell.StyleIndex = item.StyleIndex.Value;
                    }
                    else
                    {
                        if (item.Formula != null)
                        {
                            cell.CellValue = null;  //バインドだけの場合でも再計算は強制する
                        }
                        else
                        {
                            BindValue(item, cell, bindData, bindOnly);
                        }
                    }
                }
                if (!bindOnly && template.Height[rowIdx].HasValue)
                {
                    var row = GetRow(rowIndex + rowIdx);
                    row.Height = template.Height[rowIdx].Value;
                    row.CustomHeight = true;
                }
            }
            if (!bindOnly) template.MergeCells.ForEach(x => MergeCells(x.Translate(colDelta, rowDelta)));
        }

        private Dictionary<string, PropertyInfo> bindContext;

        private void BindValue(SpreadTemplate.Item item, Cell cell, object bindData, bool bindOnly)
        {
            string templateText = item.Text;
            if (bindData == null || templateText == null || !templateText.StartsWith("[") || !templateText.EndsWith("]"))
            {
                if (!bindOnly) cell.CellValue = templateText == null ? null : new CellValue(templateText);
            }
            else
            {
                string prop = templateText.Substring(1, templateText.Length - 2);
                PropertyInfo propInfo;
                if (bindContext.TryGetValue(prop, out propInfo))
                {
                    SetCellValue(cell, propInfo.GetValue(bindData, null), null, !bindOnly);
                }
                else
                {
                    var row = bindData as DataRow;
                    if (row != null && row.Table.Columns.Contains(prop))
                    {
                        SetCellValue(cell, row[prop], null, !bindOnly);
                    }
                    else
                    {
                        if (!bindOnly) cell.CellValue = new CellValue(templateText);
                    }
                }
            }
        }

        private List<SpreadsheetRange> FindMergeCells(SpreadsheetRange rangeArea)
        {
            var cells = CurrentWorksheet.GetFirstChild<MergeCells>();
            var rets = new List<SpreadsheetRange>();
            if (cells != null)
            {
                foreach (MergeCell mergeCell in cells)
                {
                    var mergeArea = new SpreadsheetRange(mergeCell.Reference.Value.ToUpper());
                    if (rangeArea.Start.ColumnIndex <= mergeArea.Start.ColumnIndex && mergeArea.Start.ColumnIndex <= rangeArea.End.ColumnIndex)
                    {
                        if (rangeArea.Start.RowIndex <= mergeArea.Start.RowIndex && mergeArea.Start.RowIndex <= rangeArea.End.RowIndex) rets.Add(mergeArea);
                    }
                }
            }
            return rets;
        }

        #endregion

        #region IDisposable メンバー

        public virtual void Dispose()
        {
            if (this.Document != null) Document.Dispose();
        }

        #endregion
    }
}
