using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Dtos;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.BLL
{
    public class EmployeeAssignmentBLL
    {
        EmployeeAssignmentDAL employeeAssignmentDAL = null;
        ExplanationsBLL explanationsBLL = null;
        public EmployeeAssignmentBLL()
        {
            employeeAssignmentDAL = new EmployeeAssignmentDAL();
            explanationsBLL = new ExplanationsBLL();
        }
        public int CreateAssignment(EmployeeAssignment employeeAssignment)
        {
            return employeeAssignmentDAL.CreateAssignment(employeeAssignment);
        }

        public int UpdateAssignment(EmployeeAssignment employeeAssignment)
        {
            return employeeAssignmentDAL.UpdateAssignment(employeeAssignment);
        }

        public List<EmployeeAssignmentViewModel> SearchAssignment(EmployeeAssignment employeeAssignment)
        {
            var employees = employeeAssignmentDAL.SearchAssignment(employeeAssignment);
            if (employees.Count > 0)
            {
                foreach (var item in employees)
                {
                    if (String.IsNullOrEmpty(item.ExplanationId))
                    {
                        item.ExplanationId = "0";
                        item.ExplanationName = "n/a";
                    }
                    else
                    {
                        item.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(item.ExplanationId)).ExplanationName;
                    }
                }
                if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
                {
                    employees = employees.Where(emp => emp.ExplanationId == employeeAssignment.ExplanationId && emp.ExplanationId != "0").ToList();
                }
            }
            return employees;
            //return employeeAssignmentDAL.SearchAssignment(employeeAssignment);
        }

        public EmployeeAssignmentViewModel GetAssignmentById(int assignmentId)
        {
            var assignment = employeeAssignmentDAL.GetAssignmentById(assignmentId);
            if (string.IsNullOrEmpty(assignment.ExplanationId))
            {
                assignment.ExplanationId = "0";
                assignment.ExplanationName = "n/a";
            }
            else
            {
                assignment.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(assignment.ExplanationId)).ExplanationName;
            }
            return assignment;
        }

        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilter(EmployeeAssignment employeeAssignment)
        {
            var employees = employeeAssignmentDAL.GetEmployeesBySearchFilter(employeeAssignment);

            if (employees.Count > 0)
            {
                int count = 1;
                foreach (var item in employees)
                {
                    if (String.IsNullOrEmpty(item.ExplanationId))
                    {
                        item.ExplanationId = "0";
                        item.ExplanationName = "n/a";
                    }
                    else
                    {
                        item.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(item.ExplanationId)).ExplanationName;
                    }
                    item.SerialNumber = count;
                    count++;
                }

                if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
                {
                    employees = employees.Where(emp => emp.ExplanationId == employeeAssignment.ExplanationId && emp.ExplanationId != "0").ToList();
                }

                this.MarkedAsRed(employees);

            }

            foreach (var item in employees)
            {
                item.EmployeeNameWithCodeRemarks += "$" + item.MarkedAsRed.ToString().ToLower();
            }
            return employees;
        }

        //public List<ForecastAssignmentViewModel> GetEmployeesForecastBySearchFilter(EmployeeAssignment employeeAssignment)
        public List<ForecastAssignmentViewModel> GetEmployeesForecastBySearchFilter(EmployeeAssignmentForecast employeeAssignment)
        {
            var employees = employeeAssignmentDAL.GetEmployeesForecastBySearchFilter(employeeAssignment);

            if (employees.Count > 0)
            {
                int count = 1;
                foreach (var item in employees)
                {
                    if (String.IsNullOrEmpty(item.ExplanationId))
                    {
                        item.ExplanationId = "0";
                        item.ExplanationName = "n/a";
                    }
                    else
                    {
                        item.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(item.ExplanationId)).ExplanationName;
                    }
                    item.SerialNumber = count;
                    count++;
                }

                //if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
                //{
                //    employees = employees.Where(emp => emp.ExplanationId == employeeAssignment.ExplanationId && emp.ExplanationId != "0").ToList();
                //}

                List<ForecastAssignmentViewModel> redMarkedForecastAssignments =  this.MarkedAsRedForForecast(employees);
                if (redMarkedForecastAssignments.Count > 0)
                {
                    foreach (var item in redMarkedForecastAssignments)
                    {
                        item.forecasts = employeeAssignmentDAL.GetForecastsByAssignmentId(item.Id);
                    }
                }

            }
            foreach (var item in employees)
            {
                item.EmployeeNameWithCodeRemarks += "$" + item.MarkedAsRed.ToString().ToLower() + "$" + item.Id; ;
            }
            return employees;
        }

        public int RemoveAssignment(int rowId)
        {
            return employeeAssignmentDAL.RemoveAssignment(rowId);
        }

        public List<EmployeeAssignmentViewModel> GetEmployeesByName(string employeeName)
        {
            var employees = employeeAssignmentDAL.GetEmployeesByName(employeeName);
            if (employees.Count > 0)
            {
                foreach (var item in employees)
                {
                    if (String.IsNullOrEmpty(item.ExplanationId))
                    {
                        item.ExplanationId = "0";
                        item.ExplanationName = "n/a";
                    }
                    else
                    {
                        item.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(item.ExplanationId)).ExplanationName;
                    }
                }
            }
            return employees;
        }

        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilterForMultipleSearch(EmployeeAssignmentDTO employeeAssignment)
        {
            List<EmployeeAssignmentViewModel> employeesWithIn = new List<EmployeeAssignmentViewModel>();
            var employees = employeeAssignmentDAL.GetEmployeesBySearchFilterForMultipleSearch(employeeAssignment);
            if (employees.Count > 0)
            {
                foreach (var item in employees)
                {
                    if (String.IsNullOrEmpty(item.ExplanationId))
                    {
                        item.ExplanationId = "0";
                        item.ExplanationName = "n/a";
                    }
                    else
                    {
                        item.ExplanationName = explanationsBLL.GetSingleExplanationByExplanationId(Int32.Parse(item.ExplanationId)).ExplanationName;
                    }
                }
                if (employeeAssignment.Explanations != null)
                {
                    if (employeeAssignment.Explanations.Length > 0)
                    {
                        foreach (var item in employeeAssignment.Explanations)
                        {
                            var employeesInTemp = employees.Where(emp => emp.ExplanationId.Contains(item) && emp.ExplanationId != "0").ToList();
                            employeesWithIn.AddRange(employeesInTemp);
                        }
                        employees = employeesWithIn;
                    }
                }
                this.MarkedAsRed(employees);
            }

            foreach (var item in employees)
            {
                item.EmployeeNameWithCodeRemarks += "$" + item.MarkedAsRed.ToString().ToLower();
            }
            return employees;
        }

        public List<EmployeeAssignmentViewModel> MarkedAsRed(List<EmployeeAssignmentViewModel> employees)
        {
            List<EmployeeAssignmentViewModel> viewModels = new List<EmployeeAssignmentViewModel>();
            List<string> names = new List<string>();

            names = (from x in employees
                     select x.EmployeeName).ToList();
            names = names.Select(n => n).Distinct().ToList();

            foreach (var name in names)
            {
                viewModels = employees.Where(emp => emp.EmployeeName == name).ToList();
                if (viewModels.Count > 1)
                {
                    foreach (var item in viewModels)
                    {
                        var singleEmployee = employees.Where(emp => emp.Id == item.Id).FirstOrDefault();
                        if (singleEmployee != null)
                        {
                            //singleEmployee.SubCode = -1;
                            if (!String.IsNullOrEmpty(singleEmployee.Remarks))
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName+"$"+singleEmployee.EmployeeName +" "+ singleEmployee.SubCode + " (" + singleEmployee.Remarks + ")";
                            }
                            else
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName+"$"+singleEmployee.EmployeeName + " " + singleEmployee.SubCode;
                            }

                        }
                    }



                    EmployeeAssignmentViewModel employeeAssignmentViewModelFirst = viewModels.Where(vm => vm.SubCode == 1).FirstOrDefault();
                    if (employeeAssignmentViewModelFirst==null)
                    {
                        continue;
                    }
                    viewModels.Remove(employeeAssignmentViewModelFirst);
                    foreach (var filteredAssignment in viewModels)
                    {
                        if (!string.IsNullOrEmpty(employeeAssignmentViewModelFirst.UnitPrice))
                        {
                            if (filteredAssignment.UnitPrice != employeeAssignmentViewModelFirst.UnitPrice)
                            {
                                employees.Where(emp => emp.Id == filteredAssignment.Id).FirstOrDefault().MarkedAsRed = true;
                            }
                        }
                    }
                }
                else
                {
                    foreach (var item in viewModels)
                    {
                        var singleEmployee = employees.Where(emp => emp.Id == item.Id).FirstOrDefault();
                        if (singleEmployee!=null)
                        {
                            //singleEmployee.SubCode = -1;
                            if (!String.IsNullOrEmpty(singleEmployee.Remarks))
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName + "$" + singleEmployee.EmployeeName + " (" + singleEmployee.Remarks + ")";
                            }
                            else
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName + "$" + singleEmployee.EmployeeName;
                            }
                            
                        }
                    }
                }
            }


            return employees;
        }

        public List<ForecastAssignmentViewModel> MarkedAsRedForForecast(List<ForecastAssignmentViewModel> employees)
        {
            List<ForecastAssignmentViewModel> viewModels = new List<ForecastAssignmentViewModel>();
            List<string> names = new List<string>();

            names = (from x in employees
                     select x.EmployeeName).ToList();
            names = names.Select(n => n).Distinct().ToList();

            foreach (var name in names)
            {
                viewModels = employees.Where(emp => emp.EmployeeName == name).ToList();
                if (viewModels.Count > 1)
                {

                    foreach (var item in viewModels)
                    {
                        var singleEmployee = employees.Where(emp => emp.Id == item.Id).FirstOrDefault();
                        if (singleEmployee != null)
                        {
                            //singleEmployee.SubCode = -1;
                            if (!String.IsNullOrEmpty(singleEmployee.Remarks))
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName+"$"+ singleEmployee.EmployeeName + " " + singleEmployee.SubCode + " (" + singleEmployee.Remarks + ")";
                            }
                            else
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName + "$" + singleEmployee.EmployeeName + " " + singleEmployee.SubCode;
                            }

                        }
                    }

                    ForecastAssignmentViewModel forecastEmployeeAssignmentViewModelFirst = viewModels.Where(vm => vm.SubCode == 1).FirstOrDefault();
                    if (forecastEmployeeAssignmentViewModelFirst == null)
                    {
                        continue;
                    }
                    viewModels.Remove(forecastEmployeeAssignmentViewModelFirst);
                    foreach (var filteredAssignment in viewModels)
                    {
                        if (filteredAssignment.UnitPrice != forecastEmployeeAssignmentViewModelFirst.UnitPrice)
                        {
                            employees.Where(emp => emp.Id == filteredAssignment.Id).FirstOrDefault().MarkedAsRed = true;

                        }
                    }
                }
                else
                {
                    foreach (var item in viewModels)
                    {
                        var singleEmployee = employees.Where(emp => emp.Id == item.Id).FirstOrDefault();
                        if (singleEmployee != null)
                        {
                            //singleEmployee.SubCode = -1;
                            if (!String.IsNullOrEmpty(singleEmployee.Remarks))
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName + "$" + singleEmployee.EmployeeName + " (" + singleEmployee.Remarks + ")";
                            }
                            else
                            {
                                singleEmployee.EmployeeNameWithCodeRemarks = singleEmployee.EmployeeName + "$" + singleEmployee.EmployeeName;
                            }

                        }
                    }
                }
            }


            return employees;
        }

        public bool CheckEmployeeName(string employeeName)
        {
            return employeeAssignmentDAL.CheckEmployeeName(employeeName);
        }
    }
}