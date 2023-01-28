using CostAllocationApp.BLL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.Controllers
{
    public class ExplanationsController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public ExplanationsController()
        {
            _sectionBLL = new SectionBLL();
        }

        // GET: Explanations
        public ActionResult CreateExplanation()
        {
            return View(new ExplanationViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}