using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.Controllers
{
    public class CommonMastersController : Controller
    {
        DepartmentBLL _departmentBll = null;
        public CommonMastersController()
        {
            _departmentBll = new DepartmentBLL();
        }
        // GET: CommonMasters
        public ActionResult CreateCommonMaster()
        {
            return View(new CommonMasterViewModel { Departments = _departmentBll.GetAllDepartments() });
        }
    }
}