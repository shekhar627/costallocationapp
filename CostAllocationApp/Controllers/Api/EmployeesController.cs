﻿using System;
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

            #region validation of inputs
            if (int.TryParse(employeeAssignmentDTO.SectionId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Section Id");

                }
                employeeAssignment.SectionId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Section Id");
            }
            if (int.TryParse(employeeAssignmentDTO.DepartmentId, out tempValue))
            {
                if (tempValue <= 0)
                {

                    return BadRequest("Invalid Department Id");
                }
                employeeAssignment.DepartmentId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Department Id");
            }
            if (int.TryParse(employeeAssignmentDTO.InchargeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Incharge Id");

                }
                employeeAssignment.InchargeId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Incharge Id");
            }
            if (int.TryParse(employeeAssignmentDTO.RoleId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Role Id");

                }
                employeeAssignment.RoleId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Role Id");
            }
            if (int.TryParse(employeeAssignmentDTO.ExplanationId, out tempValue))
            {
                if (tempValue <= 0)
                {

                    return BadRequest("Invalid Explanation Id");
                }
                employeeAssignment.ExplanationId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Explanation Id");
            }

            if (int.TryParse(employeeAssignmentDTO.CompanyId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Company Id");

                }
                employeeAssignment.CompanyId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Company Id");
            }
            if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            {
                if (tempValue < 0)
                {
                    return BadRequest("Invalid Unit Price");

                }
                employeeAssignment.UnitPrice = tempUnitPrice;
            }
            else
            {
                return BadRequest("Invalid Unit Price");
            }
            if (int.TryParse(employeeAssignmentDTO.GradeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Grade Id");

                }
                employeeAssignment.GradeId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Grade Id");
            }
            #endregion

            employeeAssignment.CreatedBy = "";
            employeeAssignment.CreatedDate = DateTime.Now;


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

            #region validation of inputs
            if (int.TryParse(employeeAssignmentDTO.SectionId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Section Id");

                }
                employeeAssignment.SectionId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Section Id");
            }
            if (int.TryParse(employeeAssignmentDTO.DepartmentId, out tempValue))
            {
                if (tempValue <= 0)
                {

                    return BadRequest("Invalid Department Id");
                }
                employeeAssignment.DepartmentId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Department Id");
            }
            if (int.TryParse(employeeAssignmentDTO.InchargeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Incharge Id");

                }
                employeeAssignment.InchargeId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Incharge Id");
            }
            if (int.TryParse(employeeAssignmentDTO.RoleId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Role Id");

                }
                employeeAssignment.RoleId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Role Id");
            }
            if (int.TryParse(employeeAssignmentDTO.ExplanationId, out tempValue))
            {
                if (tempValue <= 0)
                {

                    return BadRequest("Invalid Explanation Id");
                }
                employeeAssignment.ExplanationId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Explanation Id");
            }

            if (int.TryParse(employeeAssignmentDTO.CompanyId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Company Id");

                }
                employeeAssignment.CompanyId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Company Id");
            }
            if (decimal.TryParse(employeeAssignmentDTO.UnitPrice, out tempUnitPrice))
            {
                if (tempValue < 0)
                {
                    return BadRequest("Invalid Unit Price");

                }
                employeeAssignment.UnitPrice = tempUnitPrice;
            }
            else
            {
                return BadRequest("Invalid Unit Price");
            }
            if (int.TryParse(employeeAssignmentDTO.GradeId, out tempValue))
            {
                if (tempValue <= 0)
                {
                    return BadRequest("Invalid Grade Id");

                }
                employeeAssignment.GradeId = tempValue;
            }
            else
            {
                return BadRequest("Invalid Grade Id");
            }
            #endregion

            employeeAssignment.UpdatedBy = "";
            employeeAssignment.UpdatedDate = DateTime.Now;


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
        public IHttpActionResult SearchAssignment(string SectionId = "", string DepartmentId = "", string InchargeId = "", string RoleId = "", string ExplanationId = "", string CompanyId = "")
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
            if (int.TryParse(ExplanationId, out tempValue))
            {
                if (tempValue > 0)
                {

                    employeeAssignment.ExplanationId = tempValue;
                }

            }

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

    }
}