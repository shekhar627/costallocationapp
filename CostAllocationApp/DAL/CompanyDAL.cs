using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class CompanyDAL : DbContext
    {
        public int CreateCompany(Company company)
        {
            int result = 0;
            string query = $@"insert into Companies(Name,CreatedBy,CreatedDate,IsActive) values(@companyName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@companyName", company.CompanyName);
                cmd.Parameters.AddWithValue("@createdBy", company.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", company.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", company.IsActive);
                try
                {
                    result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                }

                return result;
            }

        }
        public List<Company> GetAllCompanies()
        {
            List<Company> companies = new List<Company>();
            string query = "";
            query = "SELECT * FROM Companies WHERE IsActive=1 ";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            Company company = new Company();
                            company.Id = Convert.ToInt32(rdr["Id"]);
                            company.CompanyName = rdr["Name"].ToString();
                            company.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            company.CreatedBy = rdr["CreatedBy"].ToString();

                            companies.Add(company);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return companies;
            }
        }
        public int RemoveCompanies(int companyIds)
        {
            int result = 0;
            string query = $@"update Companies set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", companyIds);
                try
                {
                    result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                }

                return result;
            }

        }
        public bool CheckComany(string companyName)
        {
            string query = "select * from Companies where Name=N'" + companyName + "'";
            bool result = false;
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        result = true;
                    }
                }
                catch (Exception ex)
                {

                }

                return result;
            }
        }
    }
}