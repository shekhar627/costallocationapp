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
            if (employees.Count>0)
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
               
            }
            return employees;
        }
    }
}