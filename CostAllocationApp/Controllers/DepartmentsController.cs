using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class DepartmentsController : Controller
    {
        private SectionBLL _sectionBLL = null;
        public DepartmentsController()
        {
            _sectionBLL = new SectionBLL();
        }
        // GET: Departments
        public ActionResult CreateDepartment()
        {
            return View(new DepartmentViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}