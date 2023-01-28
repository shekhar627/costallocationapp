using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class SalariesController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public SalariesController()
        {
            _sectionBLL = new SectionBLL();
        }
        // GET: Salaries
        public ActionResult CreateSalary()
        {
            return View(new SalaryViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}