using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.Controllers
{
    public class GradesController : Controller
    {
        private DepartmentBLL _departmentBLL = null;
        public GradesController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        // GET: Grades
        public ActionResult CreateGrade()
        {
            return View(new GradeViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}