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
        private DepartmentBLL _departmentBLL = null;
        public CompaniesController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        // GET: Companies
        public ActionResult CreateCompany()
        {
            return View(new CompanyViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}