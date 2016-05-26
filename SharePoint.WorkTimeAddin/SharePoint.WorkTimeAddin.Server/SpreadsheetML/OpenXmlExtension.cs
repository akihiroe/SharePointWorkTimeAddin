using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace SharePoint.WorkTimeAddin.SpreadsheetML
{
    /// <summary>
    /// SpreadsheetMLの拡張メソッド
    /// </summary>
    internal static class OpenXmlExtension
    {
        /// <summary>
        /// 複製してTに変換します。
        /// </summary>
        /// <typeparam name="T">変換する型</typeparam>
        /// <param name="value">複数元</param>
        /// <returns>複製したオブジェクト</returns>
        public static T CloneElement<T>(this OpenXmlElement value) where T:OpenXmlElement     
        {
            return value.CloneNode(true) as T;
        }

    }
}
