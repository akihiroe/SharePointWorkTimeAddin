using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LightSwitchApplication
{
    /// <summary>
    /// Download の概要の説明です
    /// </summary>
    public class SheetDownload : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            var email = context.Request.QueryString["email"];
            var year = context.Request.QueryString["year"];
            var month = context.Request.QueryString["month"];

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
                    excel.Export(temp, email, yearInt, monthInt);
                    context.Response.AddHeader("content-disposition", "attachment; filename=worktime" + monthInt.ToString("00") +".xlsx");
                    context.Response.ContentType = "application/octet-stream";
                    context.Response.WriteFile(temp);
                    context.Response.Flush();
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