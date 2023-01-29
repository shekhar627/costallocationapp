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
        private DepartmentBLL _departmentBLL = null;
        public DepartmentsController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        // GET: Departments
        public ActionResult CreateDepartment()
        {
            return View(new DepartmentViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}