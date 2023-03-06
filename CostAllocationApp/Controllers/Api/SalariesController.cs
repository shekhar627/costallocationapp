using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.Controllers.Api
{
    public class SalariesController :  ApiController
    {
        SalaryBLL salaryBLL = null;
        public SalariesController()
        {
            salaryBLL = new SalaryBLL();
        }

        //[HttpPost]
        //public IHttpActionResult CreateSalary(Salary salary)
        //{

        //    if (String.IsNullOrEmpty(salary.SalaryGrade))
        //    {
        //        return BadRequest("Grade Required");
        //    }
        //    else
        //    {
        //        salary.CreatedBy = "";
        //        salary.CreatedDate = DateTime.Now;
        //        salary.IsActive = true;


        //        //if (salaryBLL.CheckGrade(salary))
        //        //{
        //        //    return BadRequest("Data Already Exists!!!");
        //        //}
        //        //else
        //        //{
        //        //    int result = salaryBLL.CreateSalary(salary);
        //        //    if (result > 0)
        //        //    {
        //        //        return Ok("Data Saved Successfully!");
        //        //    }
        //        //    else
        //        //    {
        //        //        return BadRequest("Something Went Wrong!!!");
        //        //    }
        //        //}
        //        int result = salaryBLL.CreateSalary(salary);
        //        if (result > 0)
        //        {
        //            return Ok("Data Saved Successfully!");
        //        }
        //        else
        //        {
        //            return BadRequest("Something Went Wrong!!!");
        //        }


        //    }
        //}

        [HttpPost]
        public IHttpActionResult CreateGradeSalaryType(GradeUnitPriceType gradeUnitPriceType)
        {

            gradeUnitPriceType.CreatedBy = "";
            gradeUnitPriceType.CreatedDate = DateTime.Now;
            //salary.IsActive = true;

            var checkResult = salaryBLL.CheckGradeUnitPriceType(gradeUnitPriceType.GradeId, gradeUnitPriceType.UnitPriceTypeId, gradeUnitPriceType.DepartmentId, gradeUnitPriceType.Year);
            if (!checkResult)
            {
                int result = salaryBLL.CreateGradeUnitPriceType(gradeUnitPriceType);
                if (gradeUnitPriceType.UnitPriceTypeId == 1)
                {
                    // auto save 固定時間外単価 
                    double calculationForSecondUnitPriceType = (gradeUnitPriceType.GradeLowPoints / 160) * 1.25 * 45;
                    salaryBLL.CreateGradeUnitPriceType(new GradeUnitPriceType { GradeId = gradeUnitPriceType.GradeId, UnitPriceTypeId = 2, DepartmentId = gradeUnitPriceType.DepartmentId, Year = gradeUnitPriceType.Year, GradeLowPoints = calculationForSecondUnitPriceType, GradeHighPoints = calculationForSecondUnitPriceType, CreatedBy = "", CreatedDate = DateTime.Now });
                    // auto save 残業手当分単価 
                    double calculationForThirdUnitPriceType = gradeUnitPriceType.GradeLowPoints / 160;
                    salaryBLL.CreateGradeUnitPriceType(new GradeUnitPriceType { GradeId = gradeUnitPriceType.GradeId, UnitPriceTypeId = 3, DepartmentId = gradeUnitPriceType.DepartmentId, Year = gradeUnitPriceType.Year, GradeLowPoints = calculationForThirdUnitPriceType, GradeHighPoints = calculationForThirdUnitPriceType, CreatedBy = "", CreatedDate = DateTime.Now });

                }

                if (result > 0)
                {
                    return Ok("Data Saved Successfully!");
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Data Already Exists!!!");
            }



        }

        [HttpGet]
        public IHttpActionResult Salaries()
        {
            //List<Salary> salaries = salaryBLL.GetAllSalaryPoints();
            List<GradeSalaryTypeViewModel> salaries = salaryBLL.GetAllSalaryTypes();
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