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
            //else if (String.IsNullOrEmpty(department.SectionId.ToString()))
            //{
            //    return BadRequest("Please select section");
            //}
            else
            {
                department.CreatedBy = "";
                department.CreatedDate = DateTime.Now;
                department.IsActive = true;

                if (departmentBLL.CheckDepartment(department))
                {
                    return BadRequest("Department Already Exists!!!");
                }
                else
                {
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
        [HttpGet]
        public IHttpActionResult Departments()
        {
            List<Department> departments = departmentBLL.GetAllDepartments();
            return Ok(departments);
        }

        [HttpDelete]
        public IHttpActionResult RemoveDepartment([FromUri] string departmentIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(departmentIds))
            {
                string[] ids = departmentIds.Split(',');

                foreach (var item in ids)
                {
                    result += departmentBLL.RemoveDepartment(Convert.ToInt32(item));
                }

                if (result == ids.Length)
                {
                    return Ok("Data Removed Successfully!");
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Select Department Id!");
            }

        }

    }
}
