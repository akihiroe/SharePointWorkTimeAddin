using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LightSwitchApplication
{
    /// <summary>
    /// MailPost の概要の説明です
    /// </summary>
    public class MailPost : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            using (var serverContext = ServerApplicationContext.CreateContext())
            {
                var from = context.Request.Form["from"];
                var subject = context.Request.Form["subject"];

                using (var workspace = serverContext.Application.CreateDataWorkspace())
                {
                    var item = workspace.ApplicationData.WorkTimeSet.AddNew();
                    item.UserId = from;
                    var t = subject.Split('-');
                    if (t.Length < 2) return;
                    item.StartTime = t[0];
                    item.EndTime = t[1];
                    workspace.ApplicationData.SaveChanges();
                }
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