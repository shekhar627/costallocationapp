using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class InChargesController : Controller
    {
        // GET: InCharges
        public ActionResult CreateInCharge()
        {
            return View();
        }
    }
}