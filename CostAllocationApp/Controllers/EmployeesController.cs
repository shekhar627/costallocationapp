using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class EmployeesController : Controller
    {
        private DepartmentBLL _departmentBLL = null;
        public EmployeesController()
        {
            _departmentBLL = new DepartmentBLL();
        }

        public ActionResult CreateEmployee()
        {
            return View(new EmployeeViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
        // GET: Employees
        public ActionResult NameList()
        {
            return View(new EmployeeViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }

        public ActionResult CreateAssignment()
        {
            return View(new EmployeeViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
        public ActionResult NameRegistration()
        {
            return View(new EmployeeViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}