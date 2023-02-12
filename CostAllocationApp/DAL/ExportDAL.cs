using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Dtos;

namespace CostAllocationApp.DAL
{
    public class ExportDAL : DbContext
    {
        public List<ForecastAssignmentViewModel> AssignmentsByAllocation(int departnmentId,int explanationId)
        {
            string query = $@"select EmployeesAssignments.Id as AssignmentId, GradeSalarlyTypes.GradeId, EmployeesAssignments.SectionId,Sections.Name as SectionName,
                            EmployeesAssignments.CompanyId,Companies.Name as CompanyName
                            from EmployeesAssignments
                            left join GradeSalarlyTypes on GradeSalarlyTypes.Id = EmployeesAssignments.GradeId
                            join Companies on Companies.Id = EmployeesAssignments.CompanyId
                            join Sections on Sections.Id = EmployeesAssignments.SectionId
                            where EmployeesAssignments.DepartmentId=@departmentId and ExplanationId=@allocationId order by SectionId";

            List<ForecastAssignmentViewModel> employeeAssignments = new List<ForecastAssignmentViewModel>();

            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@departmentId",departnmentId);
                cmd.Parameters.AddWithValue("@allocationId", explanationId);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            ForecastAssignmentViewModel _forecastAssignmentViewModel = new ForecastAssignmentViewModel();
                            _forecastAssignmentViewModel.Id = Convert.ToInt32(rdr["AssignmentId"]);
                            if (rdr["GradeId"]!=DBNull.Value)
                            {
                                _forecastAssignmentViewModel.GradeId = rdr["GradeId"].ToString();
                            }
                            if (rdr["SectionId"] != DBNull.Value)
                            {
                                _forecastAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                                _forecastAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            }
                            if (rdr["CompanyId"] != DBNull.Value)
                            {
                                _forecastAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                                _forecastAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            }

                            //employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            //employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            //employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            //employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            //employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            //employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            //employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();

                            //employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(rdr["UnitPrice"]).ToString("N0");
                            //employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"] is DBNull ? "" : Convert.ToInt32(rdr["UnitPrice"]).ToString("N0");
                            //employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
                            //employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            //employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);



                            employeeAssignments.Add(_forecastAssignmentViewModel);
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return employeeAssignments;
        }

        public List<ForecastAssignmentViewModel> AssignmentsByAllocationForSection(int departnmentId, int explanationId)
        {
            string query = $@"select EmployeesAssignments.Id as AssignmentId, GradeSalarlyTypes.GradeId, EmployeesAssignments.SectionId,Sections.Name as SectionName,
                            EmployeesAssignments.CompanyId,Companies.Name as CompanyName
                            from EmployeesAssignments
                            left join GradeSalarlyTypes on GradeSalarlyTypes.Id = EmployeesAssignments.GradeId
                            join Companies on Companies.Id = EmployeesAssignments.CompanyId
                            join Sections on Sections.Id = EmployeesAssignments.SectionId
                            where DepartmentId=@departmentId and ExplanationId=@allocationId and GradeId is not null and SectionId is not null and CompanyId is not null";

            List<ForecastAssignmentViewModel> employeeAssignments = new List<ForecastAssignmentViewModel>();

            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@departmentId", departnmentId);
                cmd.Parameters.AddWithValue("@allocationId", explanationId);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            ForecastAssignmentViewModel _forecastAssignmentViewModel = new ForecastAssignmentViewModel();
                            _forecastAssignmentViewModel.Id = Convert.ToInt32(rdr["Id"]);
                            _forecastAssignmentViewModel.GradeId = rdr["GradeId"].ToString();
                            _forecastAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            _forecastAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            //employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            //employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            //employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            //employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            //employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            //employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            //employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            _forecastAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            _forecastAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            //employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(rdr["UnitPrice"]).ToString("N0");
                            //employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"] is DBNull ? "" : Convert.ToInt32(rdr["UnitPrice"]).ToString("N0");
                            //employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
                            //employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            //employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);



                            employeeAssignments.Add(_forecastAssignmentViewModel);
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return employeeAssignments;
        }

        public List<GradeSalaryType> GetGradeSalaryTypes(int gradeId,int departmentId,int year,int salaryTypeId)
        {
            string query = $@"select * from GradeSalarlyTypes where DepartmentId=@departmentId and GradeId=@gradeId and year=@year and SalaryTypeId=@salaryTypeId";
            List<GradeSalaryType> gradeSalaryTypes = new List<GradeSalaryType>();

            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@departmentId", departmentId);
                cmd.Parameters.AddWithValue("@gradeId", gradeId);
                cmd.Parameters.AddWithValue("@year", year);
                cmd.Parameters.AddWithValue("@salaryTypeId", salaryTypeId);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            GradeSalaryType gradeSalaryType = new GradeSalaryType();
                            gradeSalaryType.Id = Convert.ToInt32(rdr["Id"]);
                            gradeSalaryType.GradeId = Convert.ToInt32(rdr["GradeId"]);

                            gradeSalaryTypes.Add(gradeSalaryType);
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return gradeSalaryTypes;
        }

    }
}