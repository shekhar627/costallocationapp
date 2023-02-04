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
            string where = "";
            where += " 1=1 ";
            string query = $@"select * from EmployeesAssignments where 
                            DepartmentId=@departmentId and ExplanationId=@allocationId and GradeId is not null";

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
                            _forecastAssignmentViewModel.Id = Convert.ToInt32(rdr["Id"]);
                            _forecastAssignmentViewModel.GradeId = rdr["GradeId"].ToString();
                            //employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            //employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            //employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            //employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            //employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            //employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            //employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            //employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            //employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            //employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
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



    }
}