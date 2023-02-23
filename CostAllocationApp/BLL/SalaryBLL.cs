using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Dtos;

namespace CostAllocationApp.BLL
{
    public class SalaryBLL
    {
        SalaryDAL salaryDAL = null;
        UnitPriceTypeBLL _unitPriceTypeBLL = null;
        public SalaryBLL()
        {
            salaryDAL = new SalaryDAL();
            _unitPriceTypeBLL = new UnitPriceTypeBLL();
        }
        public int CreateSalary(Salary salary)
        {
            return salaryDAL.CreateSalary(salary);
        }
        public List<Salary> GetAllSalaryPoints()
        {
            return salaryDAL.GetAllSalaryPoints();
        }
        public List<GradeSalaryTypeViewModel> GetAllSalaryTypes()
        {
            return salaryDAL.GetAllSalaryTypes();
        }
        public List<GradeSalaryTypeViewModel> GetAllSalaries()
        {
            return salaryDAL.GetAllSalaries();
        }
        public int GetGradeId(string salaryTypeId)
        {
            return salaryDAL.GetGradeId(salaryTypeId);
        }
        public int RemoveSalary(int salaryIds)
        {
            return salaryDAL.RemoveSalary(salaryIds);
        }

        public bool CheckGrade(Salary salary)
        {
            return salaryDAL.CheckGrade(salary);
        }

        public int GetSalaryCountWithEmployeeAsignment(int gradeId)
        {
            return salaryDAL.GetSalaryCountWithEmployeeAsignment(gradeId);
        }

        public Salary GetSalaryBySalaryId(int salaryId)
        {
            return salaryDAL.GetSalaryBySalaryId(salaryId);
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

        public int CreateGradeSalaryType(GradeSalaryType gradeSalaryType)
        {
            return salaryDAL.CreateGradeSalaryType(gradeSalaryType);
        }
        public GradeSalaryType GetGradeSalaryType(int departmentId, int salaryTypeId, int year, int gradeId)
        {
            return salaryDAL.GetGradeSalaryType(departmentId, salaryTypeId, year, gradeId);
        }

        public List<int> GetSalaryTypeIdByYear(int year)
        {
            return salaryDAL.GetSalaryTypeIdByYear(year);
        }
        public SalaryMasterExportDto GetSalaryTypeWithGradeSalaryByYear(int year,int gradeId,int salaryTypeId)
        {

             return   new SalaryMasterExportDto {
                    SalaryType = _unitPriceTypeBLL.GetUnitPriceTypeById(salaryTypeId),
                    GradeSalaryTypes = salaryDAL.GetGradeSalaryTypeByYear_SalaryTypeId_GradeId(salaryTypeId, year, gradeId)
                };
            
        }
        public GradeSalaryTypeViewModel GetUnitPrice(string gradeId, string departmentId, string year)
        {
            return salaryDAL.GetUnitPrice(gradeId, departmentId, year);
        }
        public GradeSalaryTypeViewModel GetGradeSalaryTypeId(string gradeId, string departmentId)
        {
            return salaryDAL.GetGradeSalaryTypeId(gradeId, departmentId);
        }
        public bool CheckGradeSalaryType(int gradeId, int salaryTypeId, int departmentId, int year)
        {
            return salaryDAL.CheckGradeSalaryType(gradeId, salaryTypeId, departmentId, year);
        }
    }
}