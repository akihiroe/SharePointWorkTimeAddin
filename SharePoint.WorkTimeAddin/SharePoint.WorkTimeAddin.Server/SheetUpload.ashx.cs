using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LightSwitchApplication
{
    /// <summary>
    /// Upload の概要の説明です
    /// </summary>
    public class SheetUpload : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            try
            {
                var temp = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString("d") + ".xlsx";
                try
                {
                    for (int i = 0; i < context.Request.Files.Count; i++)
                    {
                        var file = context.Request.Files[i];
                        file.SaveAs(temp);
                        var excel = new ExcelManager();
                        excel.Import(temp);
                    }
                }
                finally
                {
                    System.IO.File.Delete(temp);
                }
                context.Response.ContentType = "text/plain";
                context.Response.Write("アップロードされました。");
            }
            catch (Exception ex)
            {
                context.Response.ContentType = "text/plain";
                context.Response.Write(ex.Message);
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