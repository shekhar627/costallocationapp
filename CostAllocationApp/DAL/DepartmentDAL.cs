using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class DepartmentDAL : DbContext
    {
        public int CreateDepartment(Department department)
        {
            int result = 0;
            string query = $@"insert into departments(Name,SectionId,CreatedBy,CreatedDate,IsActive) values(@departmentName,@sectionId,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@departmentName", department.DepartmentName);
                cmd.Parameters.AddWithValue("@sectionId", department.SectionId);
                cmd.Parameters.AddWithValue("@createdBy", department.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", department.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", department.IsActive);
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
    }
}