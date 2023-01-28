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
        private SectionBLL _sectionBLL = null;
        public EmployeesController()
        {
            _sectionBLL = new SectionBLL();
        }

        public ActionResult CreateEmployee()
        {
            return View(new EmployeeViewModel { Sections = _sectionBLL.GetAllSections() });
        }
        // GET: Employees
        public ActionResult NameList()
        {
            return View(new EmployeeViewModel { Sections = _sectionBLL.GetAllSections() });
        }

        public ActionResult CreateAssignment()
        {
            return View(new EmployeeViewModel { Sections = _sectionBLL.GetAllSections() });
        }
        public ActionResult NameRegistration()
        {
            return View(new EmployeeViewModel { Sections = _sectionBLL.GetAllSections() });
        }
    }
}