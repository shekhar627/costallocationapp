using System;
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
        SectionBLL sectionBLL = null;
        CompanyBLL companyBLL = new CompanyBLL();
        InChargeBLL inChargeBLL = null;
        RoleBLL roleBLL = null;
        ExplanationsBLL explanationsBLL = null;
        AllocationBLL _allocationBLL = new AllocationBLL();

        public UtilitiesController()
        {
            departmentBLL = new DepartmentBLL();
            employeeAssignmentBLL = new EmployeeAssignmentBLL();
            salaryBLL = new SalaryBLL();
            sectionBLL = new SectionBLL();
            companyBLL = new CompanyBLL();
            inChargeBLL = new InChargeBLL();
            roleBLL = new RoleBLL();
            explanationsBLL = new ExplanationsBLL();
        }


        [Route("api/utilities/DepartmentsBySection/{id}")]
        [HttpGet]
        [ActionName("DepartmentsBySection")]
        public IHttpActionResult DepartmentsBySectionId(string id)
        {
            int tempValue = 0;
            if (int.TryParse(id, out tempValue))
            {
                if (tempValue > 0)
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

        [HttpGet]
        public IHttpActionResult SearchForecastEmployee(string employeeName, string sectionId, string departmentId, string inchargeId, string roleId, string explanationId, string companyId, string status, string year)
        //public IHttpActionResult SearchForecastEmployee(EmployeeAssignmentDTO employeeAssignment)
        {
            ////employeeName = "";
            ////sectionId = "";
            ////departmentId = "";
            ////inchargeId = "";
            ////roleId = "";
            ////explanationId = "";
            ////companyId = "";
            ////status = "";
            ////year = "";

            //EmployeeAssignmentf employeeAssignment = new EmployeeAssignment();
            EmployeeAssignmentForecast employeeAssignment = new EmployeeAssignmentForecast();

            if (!string.IsNullOrEmpty(employeeName))
            {
                employeeAssignment.EmployeeName = employeeName.Trim();
            }
            else
            {
                employeeAssignment.EmployeeName = "";
            }
            //if (sectionId != null)
            //{
            //    if (sectionId.Length > 0)
            //    {
            //        string ids = "";
            //        foreach (var item in sectionId)
            //        {
            //            ids += $"{item},";
            //        }
            //        ids = ids.TrimEnd(',');

            //        //where += $" ea.SectionId in ({ids}) and ";
            //    }

            //}

            if (!string.IsNullOrEmpty(sectionId))
            {
                employeeAssignment.SectionId = sectionId;
            }
            else
            {
                employeeAssignment.SectionId = "";
            }
            if (!string.IsNullOrEmpty(departmentId))
            {
                employeeAssignment.DepartmentId = departmentId;
            }
            else
            {
                employeeAssignment.DepartmentId = "";
            }
            if (!string.IsNullOrEmpty(inchargeId))
            {
                employeeAssignment.InchargeId = inchargeId;
            }
            else
            {
                employeeAssignment.InchargeId = "";
            }
            if (!string.IsNullOrEmpty(roleId))
            {
                employeeAssignment.RoleId = roleId;
            }
            else
            {
                employeeAssignment.RoleId = "";
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
                employeeAssignment.CompanyId = companyId;
            }
            else
            {
                employeeAssignment.CompanyId = "";
            }

            if (!string.IsNullOrEmpty(status))
            {
                employeeAssignment.IsActive = status;
            }
            else
            {
                employeeAssignment.IsActive = "";
            }

            List<ForecastAssignmentViewModel> forecsatEmployeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesForecastBySearchFilter(employeeAssignment);
            //List<ForecastAssignmentViewModel> forecsatEmployeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesForecastBySearchFilter(employeeAssignment);

            return Ok(forecsatEmployeeAssignmentViewModels);
            //return Ok();

        }

        //[HttpGet]
        //public IHttpActionResult SearchForecastEmployee(EmployeeAssignmentDTO employeeAssignment)
        //{
        //    List<ForecastAssignmentViewModel> forecsatEmployeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesForecastBySearchFilter(employeeAssignment);

        //    return Ok(forecsatEmployeeAssignmentViewModels);
        //}


        [Route("api/utilities/CompareGrade/{unitPrice}")]
        [HttpGet]
        public IHttpActionResult CompareGrade(string unitPrice)
        {
            decimal tempVal = 0;
            if (decimal.TryParse(unitPrice, out tempVal))
            {
                Salary salary = salaryBLL.CompareSalary(tempVal);
                if (tempVal > 0)
                {

                    if (salary != null)
                    {
                        return Ok(salary);
                    }
                    else
                    {
                        return Ok(salary);
                        //return BadRequest("Invalid Unit Price");
                    }
                }
                else
                {
                    return Ok(salary);
                    //return BadRequest("Invalid Unit Price");
                }
            }
            else
            {
                return BadRequest("Invalid Unit Price");
            }
        }

        //[Route("api/utilities/GetUnitPrice/{unitPrice}")]
        [HttpGet]
        public IHttpActionResult GetUnitPrice(string gradeId, string departmentId, string year)
        {
            GradeSalaryTypeViewModel _salaryTypeViewModel = new GradeSalaryTypeViewModel();

            _salaryTypeViewModel = salaryBLL.GetUnitPrice(gradeId, departmentId, year);
            return Ok(_salaryTypeViewModel);
        }

        [HttpGet]
        public IHttpActionResult ViewDataByDepartment(string departmentId)
        {
            string gradeId = "";
            GradeSalaryTypeViewModel _salaryTypeViewModel = new GradeSalaryTypeViewModel();

            //_salaryTypeViewModel = salaryBLL.GetGradeSalaryTypeId(gradeId, departmentId);
            return Ok(_salaryTypeViewModel);
        }

        [HttpGet]
        public IHttpActionResult GetGradeSalaryTypeId(string gradeId, string departmentId)
        {
            GradeSalaryTypeViewModel _salaryTypeViewModel = new GradeSalaryTypeViewModel();

            _salaryTypeViewModel = salaryBLL.GetGradeSalaryTypeId(gradeId, departmentId);
            return Ok(_salaryTypeViewModel);
        }
        [Route("api/utilities/GetEmployeesByName/{employeeName}")]
        [HttpGet]
        public IHttpActionResult GetEmployeesByName(string employeeName)
        {
            if (!String.IsNullOrEmpty(employeeName))
            {
                List<EmployeeAssignmentViewModel> employeeAssignmentViewModels = employeeAssignmentBLL.GetEmployeesByName(employeeName.Trim());

                if (employeeAssignmentViewModels.Count > 0)
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

        [HttpGet]
        public IHttpActionResult SectionCount(string sectionIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(sectionIds))
            {
                int tempSectionId = 0;
                String[] _ids = sectionIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempSectionId);
                        if (tempSectionId > 0)
                        {
                            var section = sectionBLL.GetSectionBySectionId(tempSectionId);
                            if (section != null)
                            {
                                int result = sectionBLL.GetSectionCountWithEmployeeAsignment(tempSectionId);
                                //countMessage.Add(result + " rows counted for " + section.SectionName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + section.SectionName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + section.SectionName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult DepartmentCount(string departmentIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(departmentIds))
            {
                int tempDepartmentId = 0;
                String[] _ids = departmentIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempDepartmentId);
                        if (tempDepartmentId > 0)
                        {
                            var department = departmentBLL.GetDepartmentByDepartemntId(tempDepartmentId);
                            if (department != null)
                            {
                                int result = departmentBLL.GetDepartmentCountWithEmployeeAsignment(tempDepartmentId);
                                //countMessage.Add(result + " rows counted for " + department.DepartmentName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + department.DepartmentName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + department.DepartmentName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult CompanyCount(string companyIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(companyIds))
            {
                int tempCompanyId = 0;
                String[] _ids = companyIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempCompanyId);
                        if (tempCompanyId > 0)
                        {
                            var company = companyBLL.GetCompanyByCompanyId(tempCompanyId);
                            if (company != null)
                            {
                                int result = companyBLL.GetCompanyCountWithEmployeeAsignment(tempCompanyId);
                                //countMessage.Add(result + " rows counted for " + company.CompanyName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + company.CompanyName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + company.CompanyName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult InChargeCount(string inChargeIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(inChargeIds))
            {
                int tempInChargeId = 0;
                String[] _ids = inChargeIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempInChargeId);
                        if (tempInChargeId > 0)
                        {
                            var inCharge = inChargeBLL.GetInChargeByInChargeId(tempInChargeId);
                            if (inCharge != null)
                            {
                                int result = inChargeBLL.GetInChargeCountWithEmployeeAsignment(tempInChargeId);
                                //countMessage.Add(result + " rows counted for " + inCharge.InChargeName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + inCharge.InChargeName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + inCharge.InChargeName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult RoleCount(string roleIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(roleIds))
            {
                int tempRoleId = 0;
                String[] _ids = roleIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempRoleId);
                        if (tempRoleId > 0)
                        {
                            var role = roleBLL.GetRoleByRoleId(tempRoleId);
                            if (role != null)
                            {
                                int result = roleBLL.GetRoleCountWithEmployeeAsignment(tempRoleId);
                                //countMessage.Add(result + " rows counted for " + role.RoleName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + role.RoleName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + role.RoleName);

                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult ExplanationCount(string explanationIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(explanationIds))
            {
                int tempExplanationId = 0;
                String[] _ids = explanationIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempExplanationId);
                        if (tempExplanationId > 0)
                        {
                            var explanation = explanationsBLL.GetExplanationByExplanationId(tempExplanationId);
                            if (explanation != null)
                            {
                                int result = explanationsBLL.GetExplanationCountWithEmployeeAsignment(tempExplanationId);
                                countMessage.Add(result + " rows counted for " + explanation.ExplanationName);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + explanation.ExplanationName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + explanation.ExplanationName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }
            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult SalaryCount(string salaryIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(salaryIds))
            {
                int tempSalaryId = 0;
                String[] _ids = salaryIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempSalaryId);
                        if (tempSalaryId > 0)
                        {
                            var salary = salaryBLL.GetSalaryBySalaryId(tempSalaryId);
                            if (salary != null)
                            {
                                int result = salaryBLL.GetSalaryCountWithEmployeeAsignment(tempSalaryId);
                                //countMessage.Add(result + " rows counted for " + salary.SalaryGrade);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + salary.SalaryGrade);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + salary.SalaryGrade);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }

        [HttpGet]
        public IHttpActionResult AllocationCount(string allocationIds)
        {
            List<string> countMessage = new List<string>();
            if (!String.IsNullOrEmpty(allocationIds))
            {
                int tempAllocationId = 0;
                String[] _ids = allocationIds.Split(',');

                if (_ids.Length > 0)
                {
                    foreach (string item in _ids)
                    {
                        Int32.TryParse(item, out tempAllocationId);
                        if (tempAllocationId > 0)
                        {
                            var allocation = _allocationBLL.GetAllocationByAllocationId(tempAllocationId);
                            if (allocation != null)
                            {
                                int result = _allocationBLL.GetAllocationCountWithEmployeeAsignment(tempAllocationId);
                                if (result > 1)
                                {
                                    countMessage.Add(result + " projects assigned for " + allocation.AllocationName);
                                }
                                else
                                {
                                    countMessage.Add(result + " project assigned for " + allocation.AllocationName);
                                }
                            }
                        }
                    }
                    return Ok(countMessage);
                }
                else
                {
                    return BadRequest("Invalid Data");
                }


            }
            else
            {
                return BadRequest("Invalid Data");
            }

        }


        //[Route("api/utilities/DeleteAssignments/")]
        //[HttpGet]
        //[ActionName("DeleteAssignments")]
        //public IHttpActionResult DeleteAssignments()
        //{
        //    string id = 0;
        //    int tempValue = 0;
        //    if (int.TryParse(id, out tempValue))
        //    {
        //        if (tempValue > 0)
        //        {
        //            List<Department> departments = departmentBLL.GetAllDepartmentsBySectionId(sectionId: tempValue);
        //            return Ok(departments);
        //        }
        //        else
        //        {
        //            return BadRequest("Something Went Wrong!!!");
        //        }
        //    }
        //    else
        //    {
        //        return BadRequest("Something Went Wrong!!!");
        //    }


        //}

        public IHttpActionResult GetAllExplanationsByDepartmentId(int departmentId)
        {
            return Ok(explanationsBLL.GetAllExplanationsByDepartmentId(departmentId));
        }

        public IHttpActionResult GetSalaryMasterList()
        {
            GradeBLL _gradeBll = new GradeBLL();
            SalaryBLL _salaryBLL = new SalaryBLL();

            List<Grade> grades = _gradeBll.GetAllGrade();
            List<SalaryMasterExportDto> salaryMasterExportDtos = new List<SalaryMasterExportDto>();

            var salaryTypeIds = _salaryBLL.GetSalaryTypeIdByYear(2022);
            foreach (var salaryTypeId in salaryTypeIds)
            {
                foreach (var grade in grades)
                {
                    salaryMasterExportDtos.Add(_salaryBLL.GetSalaryTypeWithGradeSalaryByYear(2022, grade.Id, salaryTypeId));
                }
            }

            return Ok(salaryMasterExportDtos);
        }
        [Route("api/utilities/GetAllSalaries/")]
        [HttpGet]
        public IHttpActionResult GetAllSalaries()
        {
            List<GradeSalaryTypeViewModel> salaries = salaryBLL.GetAllSalaries();
            return Ok(salaries);

        }
        [Route("api/utilities/GetGradeId/{salaryTypeId}")]
        [HttpGet]
        public int GetGradeId(string salaryTypeId)
        {
            int gradeId = salaryBLL.GetGradeId(salaryTypeId);
            return gradeId;

        }

            
    }
}
