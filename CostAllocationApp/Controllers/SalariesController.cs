using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class SalariesController : Controller
    {
        // GET: Salaries
        public ActionResult CreateSalary()
        {
            return View();
        }
    }
}