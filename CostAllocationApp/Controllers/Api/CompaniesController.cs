using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class CompaniesController :  ApiController
    {
        CompanyBLL companyBLL = null;
        public CompaniesController()
        {
            companyBLL = new CompanyBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateCompany(Company company)
        {

            if (String.IsNullOrEmpty(company.CompanyName))
            {
                return BadRequest("Company Name Required");
            }
            else
            {
                company.CreatedBy = "";
                company.CreatedDate = DateTime.Now;
                company.IsActive = true;

                if (companyBLL.CheckComany(company.CompanyName))
                {
                    return BadRequest("Company Already Exists!!!");
                }
                else
                {
                    int result = companyBLL.CreateCompany(company);
                    if (result > 0)
                    {
                        return Ok("Data Saved Successfully!");
                    }
                    else
                    {
                        return BadRequest("Something Went Wrong!!!");
                    }
                }
            }
        }
        [HttpGet]
        public IHttpActionResult Companies()
        {
            List<Company> companies = companyBLL.GetAllCompanies();
            return Ok(companies);
        }

        [HttpDelete]
        public IHttpActionResult RemoveCompanies([FromUri] string companyIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(companyIds))
            {
                string[] ids = companyIds.Split(',');

                foreach (var item in ids)
                {
                    result += companyBLL.RemoveCompanies(Convert.ToInt32(item));
                }

                if (result == ids.Length)
                {
                    return Ok("Data Removed Successfully!");
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Select InCharge Id!");
            }

        }
    }
}