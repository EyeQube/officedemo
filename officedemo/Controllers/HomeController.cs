using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.IO;
using officedemo.Models;
using officedemo.Viewmodels;
using System.Xml;

namespace officedemo.Controllers
{
    public class HomeController : Controller
    {
        private officedemoContext db = new officedemoContext();


        public ActionResult Index()
        {
            Information information = new Information();

            List<SelectListItem> months = new List<SelectListItem>();

            months = db.Months.ToList().ConvertAll(m => new SelectListItem
            {
                Value = $"{m.Id}",
                Text = m.Name
            });
            information.Month = new SelectList(months, "Value", "Text");

            List<SelectListItem> resellers = new List<SelectListItem>();

            resellers = db.Resellers.ToList().ConvertAll(r => new SelectListItem
            {
                Value = $"{r.Id}",
                Text = r.Name
            });
            information.Reseller = new SelectList(resellers, "Value", "Text");


            System.Threading.Thread.Sleep(500);


            return View(information);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index([Bind(Include = "Sales, Costs, Order, SelectedMonthId, SelectedResellerId, Month, Reseller, Email")] Information information)
        {

                if (ModelState.IsValid)
                {
                    string reseller = db.Resellers.Where(r => r.Id == information.SelectedResellerId).First().Name;
                    string month = db.Months.Where(m => m.Id == information.SelectedMonthId).First().Name;


                    string appdataPath = Environment.ExpandEnvironmentVariables("%appdata%\\Bitoreq AB\\OfficeDemo");

                    Directory.CreateDirectory(appdataPath);
                    using (var writer = XmlWriter.Create(appdataPath + "\\salesreport.xml"))
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Reseller");
                        writer.WriteElementString("Name", reseller);
                        writer.WriteElementString("Month", month);
                        writer.WriteElementString("Date", DateTime.Now.ToShortDateString());
                        writer.WriteElementString("Sale", information.Sales.ToString());
                        writer.WriteElementString("Costs", information.Costs.ToString());
                        writer.WriteElementString("Order", information.Order.ToString());
                        writer.WriteElementString("E-mail", information.Email);
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }
            

            System.Threading.Thread.Sleep(500);

            return RedirectToAction("Index", "Home");  
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}