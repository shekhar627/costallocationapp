using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.BLL
{
    public class CompanyBLL
    {
        CompanyDAL companyDAL = null;
        public CompanyBLL()
        {
            companyDAL = new CompanyDAL();
        }
        public int CreateCompany(Company company)
        {
            return companyDAL.CreateCompany(company);
        }
        public List<Company> GetAllCompanies()
        {
            return companyDAL.GetAllCompanies();
        }
        public int RemoveCompanies(int companyIds)
        {
            return companyDAL.RemoveCompanies(companyIds);
        }
    }
}