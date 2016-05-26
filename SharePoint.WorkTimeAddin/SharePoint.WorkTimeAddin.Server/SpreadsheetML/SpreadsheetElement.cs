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
    /// OpenXmlElementを中身が同じかどうか判断する仕組みを提供します
    /// </summary>
    /// <typeparam name="T">OpenXmlElement</typeparam>
    /// <remarks>
    /// OpenXmlElementのXML生成が型が同じであれば同じ順序で生成される前提で
    /// OuterXmlを比較することで全項目の比較を行っています。
    /// これによってOpenXmlElementの派生クラスに新しい項目が追加されても中身の
    /// チェックを改訂する必要がありません。
    /// </remarks>
    internal class SpreadElement<T> where T : OpenXmlElement
    {
        public T Element { get; set; }

        public SpreadElement()
        {

        }

        public SpreadElement(T element)
        {
            if (element != null) this.Element = element.CloneElement<T>();
        }

        public override bool Equals(object obj)
        {
            if (obj == null || obj.GetType() != typeof(SpreadElement<T>)) return false;
            var target = (SpreadElement<T>)obj;
            if (Element == null && target.Element == null) return true;
            if (Element == null || Element == null) return false;
            return Element.OuterXml == target.Element.OuterXml;
        }

        public override int GetHashCode()
        {
            return Element == null ? 0 : Element.InnerXml.GetHashCode();
        }

        public T CloneElement()
        {
            return Element.CloneElement<T>();
        }
    }
}
