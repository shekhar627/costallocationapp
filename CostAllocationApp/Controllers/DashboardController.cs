using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers
{
    public class DashboardController : Controller
    {
        private DepartmentBLL _departmentBLL = null;
        public DashboardController()
        {
            _departmentBLL = new DepartmentBLL();
        }
        // GET: Dashboard
        public ActionResult Index()
        {
            return View(new DashboardViewModel { Departments = _departmentBLL.GetAllDepartments() });
        }
    }
}