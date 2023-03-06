using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class UnitPriceTypeDAL : DbContext
    {
        public int CreateUnitPriceType(SalaryType unitPriceType)
        {
            int result = 0;
            string query = $@"insert into UnitPriceTypes(SalaryTypeName,CreatedBy,CreatedDate,IsActive) values(@salaryTypeName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@salaryTypeName", unitPriceType.SalaryTypeName);
                cmd.Parameters.AddWithValue("@createdBy", unitPriceType.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", unitPriceType.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", unitPriceType.IsActive);
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

        public bool CheckUnitPriceType(string unitPriceTypeName)
        {
            string query = "select * from UnitPriceTypes where UnitPriceTypeName=N'" + unitPriceTypeName + "'";
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

        public List<SalaryType> GetAllSalaryTypes()
        {
            List<SalaryType> salaryTypes = new List<SalaryType>();
            string query = "select * from SalaryTypes";
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
                            SalaryType salaryType = new SalaryType();
                            salaryType.Id = Convert.ToInt32(rdr["Id"]);
                            salaryType.SalaryTypeName = rdr["SalaryTypeName"].ToString();
                            salaryType.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            salaryType.CreatedBy = rdr["CreatedBy"].ToString();
                            //unitPriceType.IsActive = Convert.ToBoolean(rdr["IsActive"]);

                            salaryTypes.Add(salaryType);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return salaryTypes;
            }
        }
        public List<UnitPriceType> GetAllUnitPriceTypes()
        {
            List<UnitPriceType> unitPriceTypes = new List<UnitPriceType>();
            string query = "select * from UnitPriceTypes";
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
                            UnitPriceType unitPriceType = new UnitPriceType();
                            unitPriceType.Id = Convert.ToInt32(rdr["Id"]);
                            unitPriceType.TypeName = rdr["TypeName"].ToString();
                            unitPriceType.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            unitPriceType.CreatedBy = rdr["CreatedBy"].ToString();
                            //unitPriceType.IsActive = Convert.ToBoolean(rdr["IsActive"]);

                            unitPriceTypes.Add(unitPriceType);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return unitPriceTypes;
            }
        }

        public UnitPriceType GetUnitPriceTypeById(int unitPriceTypeId)
        {
            UnitPriceType unitPriceType = new UnitPriceType();
            string query = "select * from UnitPriceTypes where id = " + unitPriceTypeId;
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
                            unitPriceType.Id = Convert.ToInt32(rdr["Id"]);
                            unitPriceType.TypeName = rdr["TypeName"].ToString();
                            unitPriceType.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            unitPriceType.CreatedBy = rdr["CreatedBy"].ToString();
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return unitPriceType;
            }
        }
    }
}