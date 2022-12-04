using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.BLL
{
    public class EmployeeAssignmentBLL
    {
        EmployeeAssignmentDAL employeeAssignmentDAL = null;
        public EmployeeAssignmentBLL()
        {
            employeeAssignmentDAL = new EmployeeAssignmentDAL();
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
            return employeeAssignmentDAL.SearchAssignment(employeeAssignment);
        }

        public EmployeeAssignmentViewModel GetAssignmentById(int assignmentId)
        {
            return employeeAssignmentDAL.GetAssignmentById(assignmentId);
        }
        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilter(EmployeeAssignment employeeAssignment)
        {
            return employeeAssignmentDAL.GetEmployeesBySearchFilter(employeeAssignment);
        }
        public int RemoveAssignment(int rowId)
        {
            return employeeAssignmentDAL.RemoveAssignment(rowId);
        }

        public List<EmployeeAssignmentViewModel> GetEmployeesByName(string employeeName)
        {
            return employeeAssignmentDAL.GetEmployeesByName(employeeName);
        }
    }
}