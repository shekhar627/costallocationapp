using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class CompaniesController : Controller
    {
        // GET: Companies
        public ActionResult CreateCompany()
        {
            return View();
        }
    }
}