using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class DepartmentBLL
    {
        DepartmentDAL departmentDAL = null;
        public DepartmentBLL()
        {
            departmentDAL = new DepartmentDAL();
        }
        public int CreateDepartment(Department department)
        {
            return departmentDAL.CreateDepartment(department);
        }
        public List<Department> GetAllDepartments()
        {
            return departmentDAL.GetAllDepartments();
        }
        public int RemoveDepartment(int departmentIds)
        {
            return departmentDAL.RemoveDepartment(departmentIds);
        }
        public List<Department> GetAllDepartmentsBySectionId(int sectionId)
        {
            return departmentDAL.GetAllDepartmentsBySectionId(sectionId);
        }
        public bool CheckDepartment(Department department)
        {
            return departmentDAL.CheckDepartment(department);
        }
        public int GetDepartmentCountWithEmployeeAsignment(int departmentId)
        {
            return departmentDAL.GetDepartmentCountWithEmployeeAsignment(departmentId);
        }
        public Department GetDepartmentByDepartemntId(int departmentId)
        {
            return departmentDAL.GetDepartmentByDepartemntId(departmentId);
        }
    }
}