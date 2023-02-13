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
        private DepartmentBLL _departmentBLL = null;
        public SalariesController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        // GET: Salaries
        public ActionResult CreateSalary()
        {
            return View(new SalaryViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
        public ActionResult SalaryMaster()
        {
            return View();
        }
    }
}