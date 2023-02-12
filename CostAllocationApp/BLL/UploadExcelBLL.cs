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
    public class UploadExcelBLL
    {
        SalaryDAL salaryDAL = null;
        SalaryBLL salaryBLL = null;
        public UploadExcelBLL()
        {
            salaryDAL = new SalaryDAL();
            salaryBLL = new SalaryBLL();
        }
        UploadExcelDAL _uploadExcelDAL = new UploadExcelDAL();
        public int GetSectionIdByName(string sectionName)
        {
            return _uploadExcelDAL.GetSectionIdByName(sectionName);
        }
        public int GetDepartmentIdByName(string departmentName)
        {
            return _uploadExcelDAL.GetDepartmentIdByName(departmentName);
        }
        public int GetInchargeIdByName(string inchargeName)
        {
            return _uploadExcelDAL.GetInchargeIdByName(inchargeName);
        }
        public int GetRoleIdByName(string roleName)
        {
            return _uploadExcelDAL.GetRoleIdByName(roleName);
        }
        public int? GetExplanationIdByName(string explanationName)
        {
            return _uploadExcelDAL.GetExplanationIdByName(explanationName);
        }
        public int GetCompanyIdByName(string companyName)
        {
            return _uploadExcelDAL.GetCompanyIdByName(companyName);
        }        
        public int GetGradeIdByUnitPrice(string unitPrice)
        {
            Salary salary = null;
            decimal tempVal = 0;
            if (decimal.TryParse(unitPrice, out tempVal))
            {
                if (tempVal > 0)
                {
                    salary = salaryBLL.CompareSalary(tempVal);                    
            }
            }
            //return _uploadExcelDAL.GetGradeIdByUnitPrice(unitPrice);
            if(salary == null)
            {
                return 0;
            }
            else
            {
                return salary.Id;
            }            
        }
        public bool IsValidMasterSetting(int masterSettingId)
        {
            bool returnValue = true;
            if (string.IsNullOrEmpty(masterSettingId.ToString()))
            {
                returnValue = false;
            }
            return returnValue;
        }
        public UploadExcel GetGradeIdByGradePoints(string gradePoints)
        {
            return _uploadExcelDAL.GetGradeIdByGradePoints(gradePoints);
        }
        public GradeSalaryType GetGradeSalaryTypeIdByGradeId(int? gradeId, int? departmentId,int year,int salaryTypeId)
        {
            return _uploadExcelDAL.GetGradeSalaryTypeIdByGradeId(gradeId, departmentId, year, salaryTypeId);
        }
    }
}