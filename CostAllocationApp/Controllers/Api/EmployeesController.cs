using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;
using CostAllocationApp.Dtos;
using CostAllocationApp.ViewModels;


namespace CostAllocationApp.Controllers.Api
{
    public class EmployeesController : ApiController
    {
        // GET: Employees
        EmployeeAssignmentBLL employeeAssignmentBLL = null;
        public EmployeesController()
        {
            employeeAssignmentBLL = new EmployeeAssignmentBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateAssignment(EmployeeAssignmentDTO employeeAssignmentDTO)
        {
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();

            int tempValue = 0;
            decimal tempUnitPrice = 0;


            if (!String.IsNullOrEmpty(employeeAssignmentDTO.EmployeeName))
            {
                var checkResult = employeeAssignmentBLL.CheckEmployeeName(employeeAssignmentDTO.EmployeeName.Trim());
                if (checkResult && employeeAssignmentDTO.SubCode == 1)
                {
                    return BadRequest("Employee Already Exists");
                }
                else
                {
                    employeeAssignment.EmployeeName = employeeAssignmentDTO.EmployeeName.Trim();
                }

            }
            else
            {
                return BadRequest("Invalid Employee Name");
            }



            if (String.IsNullOrEmpty(employeeAssignmentDTO.SectionId))
            {
                employeeAssignment.SectionId = null;
            }
            else
            {
                employeeAssignment.SectionId = Convert.ToInt32(employeeAssignmentDTO.SectionId);
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.DepartmentId))
            {
                employeeAssignment.DepartmentId = null;
            }
            else
            {
                employeeAssignment.DepartmentId = Convert.ToInt32(employeeAssignmentDTO.DepartmentId);
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.CompanyId))
            {
                employeeAssignment.CompanyId = null;
            }
            else
            {
                employeeAssignment.CompanyId = Convert.ToInt32(employeeAssignmentDTO.CompanyId);
            }

            employeeAssignment.ExplanationId = employeeAssignmentDTO.ExplanationId;

            if (String.IsNullOrEmpty(employeeAssignmentDTO.UnitPrice))
            {
                employeeAssignment.UnitPrice = null;
            }
            else
            {
                employeeAssignment.UnitPrice = Convert.ToDecimal(employeeAssignmentDTO.UnitPrice);
            }

            //if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            //{
            //    if (tempUnitPrice < 0)
            //    {
            //        return BadRequest("Invalid Unit Price");

            //    }
            //    employeeAssignment.UnitPrice = tempUnitPrice;
            //}
            //else
            //{
            //    return BadRequest("Invalid Unit Price");
            //}

            if (String.IsNullOrEmpty(employeeAssignmentDTO.GradeId))
            {
                employeeAssignment.GradeId = null;
            }
            else
            {
                employeeAssignment.GradeId = Convert.ToInt32(employeeAssignmentDTO.GradeId); ;
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.Remarks))
            {
                employeeAssignmentDTO.Remarks = "";
            }

            //employeeAssignment.ExplanationId = employeeAssignmentDTO.ExplanationId;
            employeeAssignment.CreatedBy = "";
            employeeAssignment.CreatedDate = DateTime.Now;
            employeeAssignment.IsActive = "1";
            employeeAssignment.Remarks = employeeAssignmentDTO.Remarks.Trim();
            employeeAssignment.SubCode = employeeAssignmentDTO.SubCode;


            int result = employeeAssignmentBLL.CreateAssignment(employeeAssignment);
            if (result > 0)
            {
                return Ok("Data Saved Successfully!");
            }
            else
            {
                return BadRequest("Something Went Wrong!!!");
            }
        }


        [HttpPut]
        public IHttpActionResult UpdateAssignment([FromBody]  EmployeeAssignmentDTO employeeAssignmentDTO)
        {
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();

            int tempValue = 0;
            decimal tempUnitPrice = 0;

            if (String.IsNullOrEmpty(employeeAssignmentDTO.SectionId))
            {
                employeeAssignment.SectionId = null;
            }
            else
            {
                employeeAssignment.SectionId = Convert.ToInt32(employeeAssignmentDTO.SectionId);
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.DepartmentId))
            {
                employeeAssignment.DepartmentId = null;
            }
            else
            {
                employeeAssignment.DepartmentId = Convert.ToInt32(employeeAssignmentDTO.DepartmentId);
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.CompanyId))
            {
                employeeAssignment.CompanyId = null;
            }
            else
            {
                employeeAssignment.CompanyId = Convert.ToInt32(employeeAssignmentDTO.CompanyId);
            }

            employeeAssignment.ExplanationId = employeeAssignmentDTO.ExplanationId;

            if (String.IsNullOrEmpty(employeeAssignmentDTO.UnitPrice))
            {
                employeeAssignment.UnitPrice = null;
            }
            else
            {
                employeeAssignment.UnitPrice = Convert.ToDecimal(employeeAssignmentDTO.UnitPrice);
            }

            //if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            //{
            //    if (tempUnitPrice < 0)
            //    {
            //        return BadRequest("Invalid Unit Price");

            //    }
            //    employeeAssignment.UnitPrice = tempUnitPrice;
            //}
            //else
            //{
            //    return BadRequest("Invalid Unit Price");
            //}

            if (String.IsNullOrEmpty(employeeAssignmentDTO.GradeId))
            {
                employeeAssignment.GradeId = null;
            }
            else
            {
                employeeAssignment.GradeId = Convert.ToInt32(employeeAssignmentDTO.GradeId); ;
            }

            if (String.IsNullOrEmpty(employeeAssignmentDTO.Remarks))
            {
                employeeAssignmentDTO.Remarks = "";
            }

            
            employeeAssignment.UpdatedBy = "";
            employeeAssignment.UpdatedDate = DateTime.Now;
            employeeAssignment.Id = employeeAssignmentDTO.Id;
            employeeAssignment.Remarks = employeeAssignmentDTO.Remarks.Trim();


            int result = employeeAssignmentBLL.UpdateAssignment(employeeAssignment);
            if (result > 0)
            {
                return Ok("Data Saved Successfully!");
            }
            else
            {
                return BadRequest("Something Went Wrong!!!");
            }
        }

        [HttpGet]
        public IHttpActionResult SearchAssignment(string EmployeeName="", string SectionId="", string DepartmentId="", string InchargeId="", string RoleId="", string ExplanationId="", string CompanyId="", bool Status=true)
        {

            int tempValue = 0;
            // decimal tempUnitPrice = 0;


            EmployeeAssignment employeeAssignment = new EmployeeAssignment();


            #region validation of inputs
            if (int.TryParse(SectionId, out tempValue))
            {
                if (tempValue > 0)
                {
                    employeeAssignment.SectionId = tempValue;

                }
            }
            if (int.TryParse(DepartmentId, out tempValue))
            {
                if (tempValue > 0)
                {

                    employeeAssignment.DepartmentId = tempValue;
                }
            }
            if (int.TryParse(InchargeId, out tempValue))
            {
                if (tempValue > 0)
                {
                    employeeAssignment.InchargeId = tempValue;

                }
            }
            if (int.TryParse(RoleId, out tempValue))
            {
                if (tempValue > 0)
                {
                    employeeAssignment.RoleId = tempValue;

                }
            }
            //if (int.TryParse(ExplanationId, out tempValue))
            //{
            //    if (tempValue > 0)
            //    {

            //        employeeAssignment.ExplanationId = tempValue;
            //    }

            //}

            if (int.TryParse(CompanyId, out tempValue))
            {
                if (tempValue > 0)
                {
                    employeeAssignment.CompanyId = tempValue;

                }
            }
            //if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            //{
            //    if (tempValue < 0)
            //    {
            //        return BadRequest("Invalid Unit Price");

            //    }
            //    employeeAssignment.UnitPrice = tempUnitPrice;
            //}
            //else
            //{
            //    return BadRequest("Invalid Unit Price");
            //}
            //if (int.TryParse(employeeAssignmentDTO.GradeId, out tempValue))
            //{
            //    if (tempValue <= 0)
            //    {
            //        return BadRequest("Invalid Grade Id");

            //    }
            //    employeeAssignment.GradeId = tempValue;
            //}
            //else
            //{
            //    return BadRequest("Invalid Grade Id");
            //}
            #endregion

            List<EmployeeAssignmentViewModel> employeeAssignmentViewModels = employeeAssignmentBLL.SearchAssignment(employeeAssignment);

            if (employeeAssignmentViewModels.Count > 0)
            {
                return Ok(employeeAssignmentViewModels);
            }
            else
            {
                return NotFound();
            }

        }

        [HttpPut]
        public IHttpActionResult RemoveAssignment(int id)
        {
            int result =  employeeAssignmentBLL.RemoveAssignment(id);
            if (result>0)
            {
                return Ok("Data Removed Successfully");
            }
            else
            {
                return BadRequest("Something Went Wrong");
            }
        }


    }
}