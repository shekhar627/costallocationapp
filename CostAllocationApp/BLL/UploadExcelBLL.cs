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
        public int GetExplanationIdByName(string explanationName)
        {
            return _uploadExcelDAL.GetExplanationIdByName(explanationName);
        }
        public int GetCompanyIdByName(string companyName)
        {
            return _uploadExcelDAL.GetCompanyIdByName(companyName);
        }        
        public int GetGradeIdByUnitPrice(int unitPrice)
        {
            return _uploadExcelDAL.GetGradeIdByUnitPrice(unitPrice);
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
    }
}