using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class AllocationController : Controller
    {
        // GET: Allocation
        public ActionResult CreateAllocation()
        {
            return View();
        }
    }
}