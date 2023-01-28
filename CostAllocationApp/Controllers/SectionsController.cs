using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.Controllers
{
    public class SectionsController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public SectionsController()
        {
            _sectionBLL = new SectionBLL();
        }
        public ActionResult CreateSection()
        {
            return View(new SectionViewModel { Sections = _sectionBLL.GetAllSections() });
        }

    }
}