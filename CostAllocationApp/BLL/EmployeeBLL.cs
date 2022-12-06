using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class EmployeeBLL
    {
        EmployeeDAL employeeDAL = null;
        public EmployeeBLL()
        {
            employeeDAL = new EmployeeDAL();
        }
        public int CreateEmployee(Employee employee)
        {
            return employeeDAL.CreateEmployee(employee);
        }
        public List<Employee> GetAllEmployees()
        {
            return employeeDAL.GetAllEmployees();
        }
        //public int RemoveEmployee(int employeeIds)
        //{
        //    return employeeDAL.RemoveEmployee(employeeIds);
        //}
    }
}