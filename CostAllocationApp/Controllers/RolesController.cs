using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class RolesController : Controller
    {
        // GET: Roles
        public ActionResult CreateRoles()
        {
            return View();
        }
    }
}