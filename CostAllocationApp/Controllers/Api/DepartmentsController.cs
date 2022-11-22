using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class DepartmentsController : ApiController
    {
        DepartmentBLL departmentBLL = null;
        public DepartmentsController()
        {
            departmentBLL = new DepartmentBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateDepartment(Department department)
        {

            if (String.IsNullOrEmpty(department.DepartmentName))
            {
                return BadRequest("Department Name Required");
            }
            else
            {
                department.CreatedBy = "";
                department.CreatedDate = DateTime.Now;
                department.IsActive = true;


                int result = departmentBLL.CreateDepartment(department);
                if (result > 0)
                {
                    return Ok("Data Saved Successfully!");
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }




        }
    }
}
