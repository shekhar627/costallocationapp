using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class RegistrationController : Controller
    {
        // GET: Registration
        public ActionResult Login()
        {
            return View();
        }
    }
}