using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Dtos;
using System.Globalization;


namespace CostAllocationApp.DAL
{
    public class UploadExcelDAL : DbContext
    {

        public int GetSectionIdByName(string sectionName)
        {
            int sectionId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(sectionName))
            {
                where += $" Name =N'{sectionName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM Sections
                            where {where}";

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
                                sectionId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }
            return sectionId;
        }

        public int GetDepartmentIdByName(string departmentName)
        {
            int departmentId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(departmentName))
            {
                where += $" Name =N'{departmentName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM Departments
                            where {where}";

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
                                departmentId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return departmentId;
        }

        public int GetInchargeIdByName(string inchargeName)
        {
            int inchargeId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(inchargeName))
            {
                where += $" Name =N'{inchargeName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM InCharges
                            where {where}";

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
                                inchargeId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return inchargeId;
        }

        public int GetRoleIdByName(string roleName)
        {
            int roleId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(roleName))
            {
                where += $" Name =N'{roleName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM Roles
                            where {where}";

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
                                roleId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return roleId;
        }

        public int? GetExplanationIdByName(string explanationName)
        {
            int? explanationId = null;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(explanationName))
            {
                where += $" Name =N'{explanationName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM Explanations
                            where {where}";

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
                                explanationId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return explanationId;
        }

        public int GetCompanyIdByName(string companyName)
        {
            int companyId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(companyName))
            {
                where += $" Name =N'{companyName}'";
                where += " AND IsActive=1 ";

                query = $@"select Id,Name FROM Companies
                            where {where}";

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
                                companyId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return companyId;
        }

        public int GetGradeIdByUnitPrice(int unitPrice)
        {
            

            int gradeId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(unitPrice.ToString()))
            {

                //query = "SELECT Id, GradePoints, GradeLowPoints, GradeHighPoints ";
                //query = query + "FROM Grades ";
                //query = query + "WHERE (GradeLowPoints BETWEEN  890000 AND 890000) OR(GradeHighPoints BETWEEN  890000 AND 890000)";

                where += $"(GradeLowPoints BETWEEN  {unitPrice} AND {unitPrice}) OR(GradeHighPoints BETWEEN  {unitPrice} AND {unitPrice})";
                where += " AND IsActive=1 ";

                query = $@"SELECT Id, GradePoints, GradeLowPoints, GradeHighPoints  FROM Grades
                        where {where}";

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
                                gradeId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }

            return gradeId;
        }
        public int GetGradeIdByGradePoints(string gradePoints)
        {
            int gradeId = 0;
            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(gradePoints.ToString()))
            {
                where += $"GradePoints = N'{gradePoints}' ";
                where += " AND IsActive=1 ";

                query = $@"SELECt Id,GradePoints FROM Grades 
                        where {where}";

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
                                gradeId = Convert.ToInt32(rdr["Id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        HttpContext.Current.Response.Write(ex);
                        HttpContext.Current.Response.End();
                    }
                }

            }
            return gradeId;
        }
    }
}