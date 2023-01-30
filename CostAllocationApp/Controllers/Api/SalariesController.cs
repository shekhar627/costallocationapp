using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;


namespace CostAllocationApp.Controllers.Api
{
    public class SalariesController :  ApiController
    {
        SalaryBLL salaryBLL = null;
        public SalariesController()
        {
            salaryBLL = new SalaryBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateSalary(Salary salary)
        {

            if (String.IsNullOrEmpty(salary.SalaryGrade))
            {
                return BadRequest("Grade Required");
            }
            else
            {
                salary.CreatedBy = "";
                salary.CreatedDate = DateTime.Now;
                salary.IsActive = true;


                //if (salaryBLL.CheckGrade(salary))
                //{
                //    return BadRequest("Data Already Exists!!!");
                //}
                //else
                //{
                //    int result = salaryBLL.CreateSalary(salary);
                //    if (result > 0)
                //    {
                //        return Ok("Data Saved Successfully!");
                //    }
                //    else
                //    {
                //        return BadRequest("Something Went Wrong!!!");
                //    }
                //}
                int result = salaryBLL.CreateSalary(salary);
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
        public IHttpActionResult Salaries()
        {
            List<Salary> salaries = salaryBLL.GetAllSalaryPoints();
            return Ok(salaries);
        }

        [HttpDelete]
        public IHttpActionResult RemoveSalary([FromUri] string salaryIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(salaryIds))
            {
                string[] ids = salaryIds.Split(',');

                foreach (var item in ids)
                {
                    result += salaryBLL.RemoveSalary(Convert.ToInt32(item));
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
                return BadRequest("Select Grade Points!");
            }

        }
    }
}