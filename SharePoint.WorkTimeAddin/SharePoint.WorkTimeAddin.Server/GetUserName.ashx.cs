using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LightSwitchApplication
{
    /// <summary>
    /// GetUserName の概要の説明です
    /// </summary>
    public class GetUserName : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            using (var serverContext = ServerApplicationContext.CreateContext())
            {
                context.Response.ContentType = "text/plain";
                context.Response.Write(serverContext.Application.User.Email);
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