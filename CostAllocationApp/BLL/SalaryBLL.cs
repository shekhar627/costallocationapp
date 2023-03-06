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

        public int CreateGradeUnitPriceType(GradeUnitPriceType gradeSalaryType)
        {
            return salaryDAL.CreateGradeUnitPriceType(gradeSalaryType);
        }
        public GradeUnitPriceType GetGradeUnitPriceType(int departmentId, int unitPriceType, int year, int gradeId)
        {
            return salaryDAL.GetGradeUnitPriceType(departmentId, unitPriceType, year, gradeId);
        }

        public List<int> GetUnitPriceTypeIdByYear(int year)
        {
            return salaryDAL.GetUnitPriceTypeIdByYear(year);
        }
        public SalaryMasterExportDto GetUnitPriceTypeWithGradeSalaryByYear(int year, int gradeId, int unitPriceTypeId)
        {

            return new SalaryMasterExportDto
            {
                UnitPriceType = _unitPriceTypeBLL.GetUnitPriceTypeById(unitPriceTypeId),
                GradeSalaryTypes = salaryDAL.GetGradeSalaryTypeByYear_UnitPriceTypeId_GradeId(unitPriceTypeId, year, gradeId)
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
        public bool CheckGradeUnitPriceType(int gradeId, int unitPriceTypeId, int departmentId, int year)
        {
            return salaryDAL.CheckGradeUnitPriceType(gradeId, unitPriceTypeId, departmentId, year);
        }
    }
}