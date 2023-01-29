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
        private DepartmentBLL _departmentBLL = null;
        public SectionsController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        public ActionResult CreateSection()
        {
            return View(new SectionViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

    }
}