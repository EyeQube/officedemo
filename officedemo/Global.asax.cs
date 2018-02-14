using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace officedemo
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            string appdataPath = Environment.ExpandEnvironmentVariables("%appdata%\\Bitoreq AB\\OfficeDemo");
            Directory.CreateDirectory(appdataPath);

            string fileVSContent = HttpContext.Current.Server.MapPath("~/Content");

            string filename = "återförsäljarrapportmall.xlsx";
            string fullPath = appdataPath + "/" + filename;

            if (!System.IO.File.Exists(fullPath))
            {
                System.IO.File.Copy(fileVSContent + "/" + filename, fullPath);
            }

            filename = "konsolidering.xlsx";
            fullPath = appdataPath + "/" + filename;

            if (!System.IO.File.Exists(fullPath))
            {
                System.IO.File.Copy(fileVSContent + "/" + filename, fullPath);
            }

            filename = "rapportmall.docx";
            fullPath = appdataPath + "/" + filename;

            if (!System.IO.File.Exists(fullPath))
            {
                System.IO.File.Copy(fileVSContent + "/" + filename, fullPath);
            }
        }
    }
}
