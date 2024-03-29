﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Dtos;

namespace CostAllocationApp.Controllers.Api
{
    public class UtilitiesController : ApiController
    {
        DepartmentBLL departmentBLL = null;
        EmployeeAssignmentBLL employeeAssignmentBLL = null;
        SalaryBLL salaryBLL = null;
        public UtilitiesController()
        {
            departmentBLL = new DepartmentBLL();
            employeeAssignmentBLL = new EmployeeAssignmentBLL();
            salaryBLL =new SalaryBLL();
        }


        [Route("api/utilities/DepartmentsBySection/{id}")]
        [HttpGet]
        [ActionName("DepartmentsBySection")]
        public IHttpActionResult DepartmentsBySectionId(string id)
        {
            int tempValue = 0;
            if (int.TryParse(id,out tempValue))
            {
                if (tempValue>0)
                {
                    List<Department> departments = departmentBLL.GetAllDepartmentsBySectionId(sectionId: tempValue);
                    return Ok(departments);
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Something Went Wrong!!!");
            }


        }

        [Route("api/utilities/AssignmentById/{id}")]
        [HttpGet]
        public IHttpActionResult AssignmentById(string id)
        {
            int tempValue = 0;
            if (int.TryParse(id, out tempValue))
            {
                if (tempValue > 0)
                {
                    EmployeeAssignmentViewModel employeeAssignmentViewModel = employeeAssignmentBLL.GetAssignmentById(assignmentId: tempValue);
                    return Ok(employeeAssignmentViewModel);
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Something Went Wrong!!!");
            }


        }
        [HttpGet]
        public IHttpActionResult SearchEmployee(string employeeName, string sectionId, string departmentId, string inchargeId, string roleId, string explanationId, string companyId, string status)
        {
            EmployeeAssignment employeeAssignment = new EmployeeAssignment();
            if (!string.IsNullOrEmpty(employeeName))
            {
                employeeAssignment.EmployeeName = employeeName.Trim();
            }
            else
            {
                employeeAssignment.EmployeeName = "";
            }
            if (!string.IsNullOrEmpty(sectionId))
            {
                employeeAssignment.SectionId = Convert.ToInt32(sectionId);
            }
            else
            {
                employeeAssignment.SectionId = 0;
            }
            if (!string.IsNullOrEmpty(departmentId))
            {
                employeeAssignment.DepartmentId = Convert.ToInt32(departmentId);
            }
            else
            {
                employeeAssignment.DepartmentId = 0;
            }
            if (!string.IsNullOrEmpty(inchargeId))
            {
                employeeAssignment.InchargeId = Convert.ToInt32(inchargeId);
            }
            else
            {
                employeeAssignment.InchargeId = 0;
            }
            if (!string.IsNullOrEmpty(roleId))
            {
                employeeAssignment.RoleId = Convert.ToInt32(roleId);
            }
            else
            {
                employeeAssignment.RoleId = 0;
            }
            //if (!string.IsNullOrEmpty(explanationId))
            //{
            //    employeeAssignment.ExplanationId = Convert.ToInt32(explanationId);
            //}
            //else
            //{
            //    employeeAssignment.ExplanationId = 0;
            //}

            employeeAssignment.ExplanationId = explanationId;
            if (!string.IsNullOrEmpty(companyId))
            {
                employeeAssignment.CompanyId = Convert.ToInt32(companyId);
            }
            else
            {
                employeeAssignment.CompanyId = 0;
            }

            if (!string.IsNullOrEmpty(status))
            {
                employeeAssignment.IsActive = status;
            }
            else
            {
                employeeAssignment.IsActive = "";
            }

            List<EmployeeAssignmentViewModel> employeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesBySearchFilter(employeeAssignment);

            return Ok(employeeAssignmentViewModels);

        }

        [Route("api/utilities/CompareGrade/{unitPrice}")]
        [HttpGet]
        public IHttpActionResult CompareGrade(string unitPrice)
        {
            decimal tempVal = 0;
            if (decimal.TryParse(unitPrice,out tempVal))
            {
                if (tempVal>0)
                {
                    Salary salary = salaryBLL.CompareSalary(tempVal);
                    if (salary!=null)
                    {
                        return Ok(salary);
                    }
                    else
                    {
                        return BadRequest("Invalid Unit Price");
                    }
                }
                else
                {
                    return BadRequest("Invalid Unit Price");
                }
            }
            else
            {
                return BadRequest("Invalid Unit Price");
            }
        }

        [Route("api/utilities/GetEmployeesByName/{employeeName}")]
        [HttpGet]
        public IHttpActionResult GetEmployeesByName(string employeeName)
        {
            if (!String.IsNullOrEmpty(employeeName))
            {
                List<EmployeeAssignmentViewModel> employeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesByName(employeeName.Trim());

                if (employeeAssignmentViewModels.Count>0)
                {
                    return Ok(employeeAssignmentViewModels);
                }
                else
                {
                    return NotFound();
                }
            }
            else
            {
                return NotFound();
            }
            
        }

        [HttpPost]
        public IHttpActionResult SearchMultipleEmployee(EmployeeAssignmentDTO employeeAssignment)
        {
            List<EmployeeAssignmentViewModel> employeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesBySearchFilterForMultipleSearch(employeeAssignment);

            return Ok(employeeAssignmentViewModels);
        }

    }
}
