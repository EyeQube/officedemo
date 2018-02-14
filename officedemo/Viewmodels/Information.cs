using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace officedemo.Viewmodels
{
    public class Information
    {
        public bool DisplayRobot { get; set; }  

        [Display(Name = "Försäljning(kkr)")]
        public int Sales { get; set; }

        [Display(Name = "Order(kkr)")]
        public int Order { get; set; }

        [Display(Name = "Kostnader(kkr)")]
        public int Costs { get; set; }

        [Display(Name = "Datum")]
        public DateTime Date { get; set; }

        public int? SelectedMonthId { get; set; }
        [Display(Name = "Månad")]
        public IEnumerable<SelectListItem> Month { get; set; }

        public int? SelectedResellerId { get; set; }
        [Display(Name = "Återförsäljare")]
        public IEnumerable<SelectListItem> Reseller { get; set; }

        [Display(Name = "Mottagare (e-post)")]
        public string Email { get; set; }

    }
}