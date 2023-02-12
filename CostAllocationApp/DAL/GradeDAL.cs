using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class GradeDAL: DbContext
    {
        public int CreateGrade(Grade grade)
        {
            int result = 0;
            string query = $@"insert into Grades(GradeName,CreatedBy,CreatedDate) values(@gradeName,@createdBy,@createdDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@gradeName", grade.GradeName);
                cmd.Parameters.AddWithValue("@createdBy", grade.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", grade.CreatedDate);
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

        public bool CheckGrade(string gradeName)
        {
            string query = "select * from Grades where GradeName=N'" + gradeName + "'";
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

        public List<Grade> GetAllGrade()
        {
            List<Grade> grades = new List<Grade>();
            string query = "select * from Grades;";
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
                            Grade grade = new Grade();
                            grade.Id = Convert.ToInt32(rdr["Id"]);
                            grade.GradeName= rdr["GradeName"].ToString();
                            grade.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            //section.CreatedBy = rdr["CreatedBy"].ToString();

                            grades.Add(grade);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return grades;
            }
        }
    }
}