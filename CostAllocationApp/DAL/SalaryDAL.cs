using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class SalaryDAL : DbContext
    {
        public int CreateSalary(Salary salary)
        {
            int result = 0;
            string query = $@"insert into Grades(GradePoints,GradeLowPoints,GradeHighPoints,CreatedBy,CreatedDate,IsActive) values(@gradePoints,@gradeLowPoints,@gradeHighPoints,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@gradePoints", salary.SalaryGrade);
                cmd.Parameters.AddWithValue("@gradeLowPoints", salary.SalaryLowPoint);
                cmd.Parameters.AddWithValue("@gradeHighPoints", salary.SalaryHighPoint);
                cmd.Parameters.AddWithValue("@createdBy", salary.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", salary.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", salary.IsActive);
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
        public List<Salary> GetAllSalaryPoints()
        {
            List<Salary> salaries = new List<Salary>();
            string query = "";
            query = "SELECT * FROM Grades WHERE IsActive=1 ";
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
                            Salary salary = new Salary();
                            salary.Id = Convert.ToInt32(rdr["Id"]);
                            salary.SalaryGrade = rdr["GradePoints"].ToString();
                            salary.SalaryLowPoint = Convert.ToDecimal(rdr["GradeLowPoints"]);
                            salary.SalaryHighPoint = Convert.ToDecimal(rdr["GradeHighPoints"]);
                            salary.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            salary.CreatedBy = rdr["CreatedBy"].ToString();

                            salaries.Add(salary);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return salaries;
            }
        }
        public int RemoveSalary(int salaryIds)
        {
            int result = 0;
            string query = $@"update Grades set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", salaryIds);
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