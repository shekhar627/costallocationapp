﻿using System;
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
        public UploadExcel GetGradeIdByGradePoints(string gradePoints)
        {
            UploadExcel _salary = new UploadExcel();

            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(gradePoints.ToString()))
            {
                where += $"GradeName = N'{gradePoints}' ";
                //where += " AND IsActive=1 ";

                query = $@"SELECt Id,GradeName FROM Grades 
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
                                _salary.GradeId = Convert.ToInt32(rdr["Id"]);
                                _salary.GradeName = rdr["GradeName"].ToString();                                                                
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
            return _salary;
        }
        public GradeUnitPriceType GetGradeUnitPriceType(int? gradeId,int? departmentId,int year,int unitPriceTypeId)
        {
            GradeUnitPriceType _salaryType = new GradeUnitPriceType();

            string where = "";
            string query = "";
            if (!string.IsNullOrEmpty(gradeId.ToString()) && !string.IsNullOrEmpty(departmentId.ToString()))
            {
                where += $"GradeId = {gradeId} and DepartmentId = {departmentId} and Year={year} and UnitPriceTypeId={unitPriceTypeId}";

                query = $@"SELECt Id,GradeId,GradeLowPoints FROM GradeUnitPriceTypes 
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
                                _salaryType.Id = Convert.ToInt32(rdr["Id"]);
                                _salaryType.GradeId = Convert.ToInt32(rdr["GradeId"]);
                                _salaryType.GradeLowPoints = Convert.ToSingle(rdr["GradeLowPoints"]);
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
            return _salaryType;
        }

    }
}