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
        public List<Department> GetAllDepartments()
        {
            List<Department> departments = new List<Department>();
            string query = "";
            query = "SELECT dpt.*,sc.Name as SectionName ";
            query = query + "FROM Departments dpt ";
            query = query + "    INNER JOIN Sections sc ON dpt.SectionId = sc.Id ";
            query = query + "WHERE dpt.isactive=1 ";
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
                            Department department = new Department();
                            department.Id = Convert.ToInt32(rdr["Id"]);
                            department.DepartmentName = rdr["Name"].ToString();
                            department.SectionName = rdr["SectionName"].ToString();
                            department.SectionId = Convert.ToInt32(rdr["SectionId"]);
                            department.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            department.CreatedBy = rdr["CreatedBy"].ToString();

                            departments.Add(department);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return departments;
            }
        }
        public int RemoveDepartment(int departmentId)
        {
            int result = 0;
            string query = $@"update Departments set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", departmentId);
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