using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class CompaniesController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public CompaniesController()
        {
            _sectionBLL = new SectionBLL();
        }
        // GET: Companies
        public ActionResult CreateCompany()
        {
            return View(new CompanyViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}