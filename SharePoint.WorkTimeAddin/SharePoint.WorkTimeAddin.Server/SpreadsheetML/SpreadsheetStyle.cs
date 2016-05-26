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
    /// スプレッドシートのスタイルを表現します。
    /// </summary>
    public class SpreadsheetStyle
    {
        internal SpreadElement<Font> fontElement;
        internal SpreadElement<Fill> fillElement;
        internal SpreadElement<Border> borderElement;
        internal SpreadElement<Alignment> alignmentElement;
        internal SpreadElement<NumberingFormat> numberingFormatElement;

        /// <summary>
        /// Fontスタイル
        /// </summary>
        public Font Font { get { return fontElement.Element; } }

        /// <summary>
        /// PatternFillスタイル
        /// </summary>
        public PatternFill PatternFill { get { return fillElement.Element.PatternFill; } }

        /// <summary>
        /// Borderスタイル
        /// </summary>
        public Border Border { get { return borderElement.Element; } }

        /// <summary>
        /// NumberingFormatスタイル
        /// </summary>
        public NumberingFormat NumberingFormat { get { return numberingFormatElement.Element; } }

        /// <summary>
        /// Alignmentスタイル
        /// </summary>
        public Alignment Alignment
        {
            get
            {
                if (alignmentElement == null) alignmentElement = new SpreadElement<Alignment>(new Alignment());
                return alignmentElement.Element;
            }
        }

        internal SpreadsheetStyle() { }

        private string GetRgb(string rgb)
        {
            if (rgb == null) return null;
            if (rgb.Length == 8) return rgb;
            if (rgb.Length == 6) return "FF" + rgb;
            throw new ArgumentException("RGB指定は６または8文字の16進数形式で指定してください("+rgb+")");
        }

        /// <summary>
        /// 文字列の色を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetColor(System.Drawing.Color color)
        {
            return SetColor(color.ToArgb().ToString("X8"));
        }

        /// <summary>
        /// 文字列の色を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetColor(string rgb)
        {
            Font.Color.Rgb = GetRgb(rgb);
            Font.Color.Theme = null;
            return this;
        }

        /// <summary>
        /// 背景色を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBackgroundColor(System.Drawing.Color color)
        {
            return SetBackgroundColor(color.ToArgb().ToString("X8"));
        }

        /// <summary>
        /// 背景色を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBackgroundColor(string rgb)
        {
            PatternFill.PatternType = new EnumValue<PatternValues>(PatternValues.Solid);
            if (PatternFill.ForegroundColor == null) PatternFill.ForegroundColor = new ForegroundColor();
            PatternFill.ForegroundColor.Rgb = GetRgb(rgb);
            PatternFill.ForegroundColor.Theme = null;
            if (PatternFill.BackgroundColor == null) PatternFill.BackgroundColor = new BackgroundColor();
            PatternFill.BackgroundColor.Rgb = GetRgb(rgb);
            PatternFill.BackgroundColor.Theme = null;
            return this;
        }

        /// <summary>
        /// 文字をイタリックにするか指定します。
        /// </summary>
        /// <param name="on">true：有効、false：無効</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetItalic(bool on = true)
        {
            Font.Italic = on ? new Italic() : null;
            return this;
        }

        /// <summary>
        /// 文字をイタリックにするか指定します。
        /// </summary>
        /// <param name="on">true：有効、false：無効</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBold(bool on = true)
        {
            Font.Bold = on ? new Bold() : null;
            return this;
        }

        /// <summary>
        /// 文字に下線を付与するか指定します。
        /// </summary>
        /// <param name="on">true：有効、false：無効</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetUnderline(bool on = true)
        {
            Font.Underline = on ? new Underline() : null;
            return this;
        }

        /// <summary>
        /// 文字に下線を付与するか指定します。
        /// </summary>
        /// <param name="on">true：有効、false：無効</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetWrapText(bool on = true)
        {
            Alignment.WrapText = new BooleanValue(on);
            return this;
        }

        /// <summary>
        /// 水平方向のアライメントを指定します。
        /// </summary>
        /// <param name="value">アライメント</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetHorizontalAlignment(HorizontalAlignmentValues value)
        {
            Alignment.Horizontal = value;
            return this;
        }

        /// <summary>
        /// 水直方向のアライメントを指定します。
        /// </summary>
        /// <param name="value">アライメント</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetVerticalAlignment(VerticalAlignmentValues value)
        {
            Alignment.Vertical = value;
            return this;
        }

        /// <summary>
        /// 日付型のフォーマットを指定します。
        /// </summary>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetDateTimeFormat()
        {
            NumberingFormat.NumberFormatId = 14;
            return this;
        }

        /// <summary>
        /// フォーマット番号を指定します。
        /// </summary>
        /// <param name="formatId"></param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetFormatCode(uint formatId)
        {
            NumberingFormat.NumberFormatId = formatId;
            return this;
        }

        /// <summary>
        /// フォーマットコードを指定します。
        /// </summary>
        /// <param name="formatCode"></param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetFormatCode(string formatCode)
        {
            NumberingFormat.FormatCode = formatCode;
            return this;
        }

        /// <summary>
        /// 文字のサイズを指定します。
        /// </summary>
        /// <param name="fontsize"></param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetFontSize(double fontsize)
        {
            Font.FontSize.Val= fontsize;
            return this;
        }

        /// <summary>
        /// 罫線（上）を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderTop(System.Drawing.Color color, BorderStyleValues style)
        {
            return SetBorderTop(color.ToArgb().ToString("X8"), style);
        }

        /// <summary>
        /// 罫線（上）を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderTop(string rgb, BorderStyleValues style)
        {
            Border.TopBorder = rgb == null ? null : new TopBorder();
            return SetBorder(Border.TopBorder, rgb, style);
        }

        /// <summary>
        /// 罫線（下）を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderBottom(System.Drawing.Color color, BorderStyleValues style)
        {
            return SetBorderBottom(color.ToArgb().ToString("X8"), style);
        }

        /// <summary>
        /// 罫線（下）を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderBottom(string rgb, BorderStyleValues style)
        {
            Border.BottomBorder = rgb == null ? null : new BottomBorder();
            return SetBorder(Border.BottomBorder, rgb, style);
        }


        /// <summary>
        /// 罫線（左）を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderLeft(System.Drawing.Color color, BorderStyleValues style)
        {
            return SetBorderLeft(color.ToArgb().ToString("X8"), style);
        }


        /// <summary>
        /// 罫線（左）を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderLeft(string rgb, BorderStyleValues style)
        {
            Border.LeftBorder = rgb == null ? null : new LeftBorder();
            return SetBorder(Border.LeftBorder, rgb, style);
        }

        /// <summary>
        /// 罫線（右）を指定します。
        /// </summary>
        /// <param name="color">色</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderRight(System.Drawing.Color color, BorderStyleValues style)
        {
            return SetBorderRight(color.ToArgb().ToString("X8"), style);
        }

        /// <summary>
        /// 罫線（右）を指定します。
        /// </summary>
        /// <param name="rgb">色（RGB形式）</param>
        /// <param name="style">罫線スタイル</param>
        /// <returns>スタイル（メソッドチェーン用）</returns>
        public SpreadsheetStyle SetBorderRight(string rgb, BorderStyleValues style)
        {
            Border.RightBorder = rgb == null ? null : new RightBorder();
            return SetBorder(Border.RightBorder, rgb, style);
        }

        internal SpreadsheetStyle SetBorder(BorderPropertiesType item, string rgb, BorderStyleValues style)
        {
            if (rgb != null)
            {
                item.Color = new Color() { Rgb = GetRgb(rgb) };
                item.Style = style;
            }
            return this;
        }


    }
}
