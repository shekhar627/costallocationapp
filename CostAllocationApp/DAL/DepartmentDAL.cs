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
            string query = $@"insert into departments(Name,CreatedBy,CreatedDate,IsActive) values(@departmentName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@departmentName", department.DepartmentName);
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
            query = "SELECT dpt.*";
            query = query + "FROM Departments dpt ";
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

        public List<Department> GetAllDepartmentsBySectionId(int sectionId)
        {
            List<Department> departments = new List<Department>();
            string query = "";
            query = "SELECT dpt.*,sc.Name as SectionName ";
            query = query + "FROM Departments dpt ";
            query = query + "    INNER JOIN Sections sc ON dpt.SectionId = sc.Id ";
            query = query + "WHERE dpt.isactive=1 and SectionId="+sectionId;
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

        public bool CheckDepartment(Department department)
        {
            string query = "select * from Departments  where Name=N'" + department.DepartmentName + "' and SectionId="+department.SectionId;
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

        public int GetDepartmentCountWithEmployeeAsignment(int departmentId)
        {
            string query = "select * from EmployeesAssignments where DepartmentId=" + departmentId;
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

        public Department GetDepartmentByDepartemntId(int departmentId)
        {
            Department department = null;
            string query = "select * from Departments where Id = " + departmentId;
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
                            department = new Department();
                            department.DepartmentName = rdr["Name"].ToString();
                            department.Id = Convert.ToInt32(rdr["Id"]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    department = null;
                }

                return department;
            }
        }
    }
}