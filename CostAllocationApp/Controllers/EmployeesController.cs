using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class EmployeesController : Controller
    {
        public ActionResult CreateEmployee()
        {
            return View();
        }
        // GET: Employees
        public ActionResult NameList()
        {
            return View();
        }
    }
}