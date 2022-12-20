using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class RoleDAL : DbContext
    {
        public int CreateRole(Role role)
        {
            int result = 0;
            string query = $@"insert into Roles(Name,CreatedBy,CreatedDate,IsActive) values(@roleName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@roleName", role.RoleName);
                cmd.Parameters.AddWithValue("@createdBy", role.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", role.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", role.IsActive);
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
        public List<Role> GetAllRoles()
        {
            List<Role> roles = new List<Role>();
            string query = "";
            query = "SELECT * FROM Roles WHERE IsActive=1 ";
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
                            Role role = new Role();
                            role.Id = Convert.ToInt32(rdr["Id"]);
                            role.RoleName = rdr["Name"].ToString();
                            role.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            role.CreatedBy = rdr["CreatedBy"].ToString();

                            roles.Add(role);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return roles;
            }
        }
        public int RemoveRoles(int roleIds)
        {
            int result = 0;
            string query = $@"update Roles set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", roleIds);
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

        public bool CheckRole(string roleName)
        {
            string query = "select * from Roles where Name=N'" + roleName + "'";
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