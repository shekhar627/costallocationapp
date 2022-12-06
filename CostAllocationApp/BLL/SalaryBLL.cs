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

        public Salary CompareSalary(decimal unitPrice)
        {
            Salary salary = null;
            List<Salary> salaries =salaryDAL.GetAllSalaryPoints();
            if (salaries.Count>0)
            {
                foreach (var item in salaries)
                {
                    bool gradePoint = this.BetweenRanges(item.SalaryLowPoint,item.SalaryHighPoint,unitPrice);
                    if (gradePoint)
                    {
                        salary = item;
                        break;
                    }
                }
            }
            return salary;
        }


        private bool BetweenRanges(decimal a, decimal b, decimal number)
        {
            return (a <= number && number <= b);
        }
    }
}