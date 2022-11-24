using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.BLL
{
    public class SalaryBLL
    {
        SalaryDAL salaryDAL = null;
        public SalaryBLL()
        {
            salaryDAL = new SalaryDAL();
        }
        public int CreateSalary(Salary salary)
        {
            return salaryDAL.CreateSalary(salary);
        }
        public List<Salary> GetAllSalaryPoints()
        {
            return salaryDAL.GetAllSalaryPoints();
        }
        public int RemoveSalary(int salaryIds)
        {
            return salaryDAL.RemoveSalary(salaryIds);
        }
    }
}