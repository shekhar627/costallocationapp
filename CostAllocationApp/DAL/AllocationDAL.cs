using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class AllocationDAL : DbContext
    {
        public int CreateAllocation(Allocation _allocation)
        {
            int result = 0;
            string query = $@"insert into Allocations(Name,CreatedBy,CreatedDate,IsActive) values(@allocationName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@allocationName", _allocation.AllocationName);
                cmd.Parameters.AddWithValue("@createdBy", _allocation.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", _allocation.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", _allocation.IsActive);
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
        public List<Allocation> GetAllocationList()
        {
            List<Allocation> _allocations = new List<Allocation>();
            string query = "select * from Allocations where isactive=1";
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
                            Allocation _allocation = new Allocation();
                            _allocation.Id = Convert.ToInt32(rdr["Id"]);
                            _allocation.AllocationName = rdr["Name"].ToString();
                            _allocation.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            _allocation.CreatedBy = rdr["CreatedBy"].ToString();

                            _allocations.Add(_allocation);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return _allocations;
            }
        }

        public int RemoveAllocation(int allocationId)
        {
            int result = 0;
            string query = $@"update Allocations set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", allocationId);
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

        public bool CheckAllocation(string allocationName)
        {
            string query = "select * from Sections where Name=N'" + allocationName + "'";
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

        public int GetAllocationCountWithEmployeeAsignment(int allocationId)
        {
            string query = "select * from EmployeesAssignments where SectionId=" + allocationId;
            int result = 0;
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
                            result++;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return result;
        }

        public Allocation GetAllocationByAllocationId(int allocationId)
        {
            Allocation _allocation = null;
            string query = "select * from Sections where Id = " + allocationId;
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
                        while (rdr.Read())
                        {
                            _allocation = new Allocation();
                            _allocation.AllocationName = rdr["Name"].ToString();
                            _allocation.Id = Convert.ToInt32(rdr["Id"]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _allocation = null;
                }

                return _allocation;
            }
        }
    }
}