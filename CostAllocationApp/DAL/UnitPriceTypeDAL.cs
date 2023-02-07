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
        public int CreateUnitPriceType(UnitPriceType unitPriceType)
        {
            int result = 0;
            string query = $@"insert into UnitPriceTypes(UnitPriceTypeName,CreatedBy,CreatedDate,IsActive) values(@unitPriceTypeName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@unitPriceTypeName", unitPriceType.UnitPriceTypeName);
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

        public List<UnitPriceType> GetAllUnitPriceTypes()
        {
            List<UnitPriceType> unitPriceTypes = new List<UnitPriceType>();
            string query = "select * from UnitPriceTypes where isactive=1";
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
                            unitPriceType.UnitPriceTypeName = rdr["UnitPriceTypeName"].ToString();
                            unitPriceType.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            unitPriceType.CreatedBy = rdr["CreatedBy"].ToString();
                            unitPriceType.IsActive = Convert.ToBoolean(rdr["IsActive"]);

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
    }
}