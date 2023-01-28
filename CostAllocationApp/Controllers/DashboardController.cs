using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class DashboardController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public DashboardController()
        {
            _sectionBLL = new SectionBLL();
        }
        // GET: Dashboard
        public ActionResult Index()
        {
            return View(new DashboardViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}