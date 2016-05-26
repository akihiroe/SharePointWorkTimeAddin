using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LightSwitchApplication
{
    /// <summary>
    /// SubmitWorkTime の概要の説明です
    /// </summary>
    public class SubmitWorkTime : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            var email = context.Request.Form["email"];
            var year = context.Request.Form["year"];
            var month = context.Request.Form["month"];

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(year) || string.IsNullOrEmpty(month))
            {
                context.Response.StatusCode = 400;
                return;
            }
            int yearInt;
            int monthInt;
            if (int.TryParse(month, out monthInt) && int.TryParse(year, out yearInt))
            {
                var temp = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString("d") + ".xlsx";
                try
                {
                    var excel = new ExcelManager();
                    System.IO.File.Copy(context.Server.MapPath("~/WorkTime_Template.xlsx"), temp);
                    excel.Submit(temp, email, yearInt, monthInt);
                    context.Response.ContentType = "text/plain";
                    context.Response.Write(string.Format("{0}月度の作業時間を送信しました", month));
                }
                catch (Exception ex)
                {
                    context.Response.ContentType = "text/plain";
                    context.Response.Write(ex.Message);
                }
                finally
                {
                    System.IO.File.Delete(temp);
                }

            }
            else
            {
                context.Response.StatusCode = 400;
                return;
            }

        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}