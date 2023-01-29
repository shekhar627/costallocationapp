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
        private DepartmentBLL _departmentBLL = null;
        public ExplanationsController()
        {
            _departmentBLL = new DepartmentBLL();
        }

        // GET: Explanations
        public ActionResult CreateExplanation()
        {
            return View(new ExplanationViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}