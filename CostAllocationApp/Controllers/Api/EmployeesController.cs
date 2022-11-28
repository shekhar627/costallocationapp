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
    public class EmployeesController : ApiController
    {
        EmployeeBLL employeeBLL = null;
        public EmployeesController()
        {
            employeeBLL = new EmployeeBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateEmployee(Employee employee)
        {

            if (String.IsNullOrEmpty(employee.FirstName))
            {
                return BadRequest("Employee Name Required");
            }
            else
            {
                employee.CreatedBy = "";
                employee.CreatedDate = DateTime.Now;
                employee.IsActive = true;


                int result = employeeBLL.CreateEmployee(employee);
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
        [HttpGet]
        public IHttpActionResult Employees()
        {
            List<Employee> employees = employeeBLL.GetAllEmployees();
            return Ok(employees);
        }

        [HttpDelete]
        public IHttpActionResult RemoveEmployee([FromUri] string employeeIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(employeeIds))
            {
                string[] ids = employeeIds.Split(',');

                foreach (var item in ids)
                {
                    result += employeeBLL.RemoveEmployee(Convert.ToInt32(item));
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
                return BadRequest("Select employee!");
            }

        }

    }
}