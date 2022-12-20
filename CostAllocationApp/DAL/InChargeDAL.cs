using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class InChargeDAL : DbContext
    {
        public int CreateInCharge(InCharge inCharge)
        {
            int result = 0;
            string query = $@"insert into InCharges(Name,CreatedBy,CreatedDate,IsActive) values(@inChargeName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@inChargeName", inCharge.InChargeName);
                cmd.Parameters.AddWithValue("@createdBy", inCharge.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", inCharge.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", inCharge.IsActive);
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
        public List<InCharge> GetAllInCharges()
        {
            List<InCharge> inCharges = new List<InCharge>();
            string query = "";
            query = "SELECT * FROM InCharges WHERE IsActive=1 ";
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
                            InCharge inCharge = new InCharge();
                            inCharge.Id = Convert.ToInt32(rdr["Id"]);
                            inCharge.InChargeName = rdr["Name"].ToString();
                            inCharge.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            inCharge.CreatedBy = rdr["CreatedBy"].ToString();

                            inCharges.Add(inCharge);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return inCharges;
            }
        }
        public int RemoveInCharge(int inChargeIds)
        {
            int result = 0;
            string query = $@"update InCharges set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", inChargeIds);
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
        public bool CheckInCharge(string incharegeName)
        {
            string query = "select * from InCharges where Name=N'" + incharegeName + "'";
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