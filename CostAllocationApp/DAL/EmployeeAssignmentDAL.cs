using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;
using CostAllocationApp.ViewModels;

namespace CostAllocationApp.DAL
{
    public class EmployeeAssignmentDAL : DbContext
    {
        public int CreateAssignment(EmployeeAssignment employeeAssignment)
        {
            int result = 0;
            string query = $@"insert into EmployeesAssignments(SectionId,DepartmentId,InChargeId,RoleId,ExplanationId,CompanyId,UnitPrice,GradeId,CreatedBy,CreatedDate) values(@sectionId,@departmentId,@inChargeId,@roleId,@explanationId,@companyId,@unitPrice,@gradeId,@createdBy,@createdDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@sectionId", employeeAssignment.SectionId);
                cmd.Parameters.AddWithValue("@departmentId", employeeAssignment.DepartmentId);
                cmd.Parameters.AddWithValue("@inChargeId", employeeAssignment.InchargeId);
                cmd.Parameters.AddWithValue("@roleId", employeeAssignment.RoleId);
                cmd.Parameters.AddWithValue("@explanationId", employeeAssignment.ExplanationId);
                cmd.Parameters.AddWithValue("@companyId", employeeAssignment.CompanyId);
                cmd.Parameters.AddWithValue("@unitPrice", employeeAssignment.UnitPrice);
                cmd.Parameters.AddWithValue("@gradeId", employeeAssignment.GradeId);
                cmd.Parameters.AddWithValue("@createdBy", employeeAssignment.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", employeeAssignment.CreatedDate);

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

        public int UpdateAssignment(EmployeeAssignment employeeAssignment)
        {
            int result = 0;
            string query = $@"update EmployeesAssignments set  SectionId=@sectionId,DepartmentId=@departmentId,InChargeId=@inChargeId,RoleId=@roleId,ExplanationId=@explanationId,CompanyId=@companyId,UnitPrice=@unitPrice,GradeId=@gradeId,UpdatedBy=@updatedBy,UpdatedDate=@updatedDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@sectionId", employeeAssignment.SectionId);
                cmd.Parameters.AddWithValue("@departmentId", employeeAssignment.DepartmentId);
                cmd.Parameters.AddWithValue("@inChargeId", employeeAssignment.InchargeId);
                cmd.Parameters.AddWithValue("@roleId", employeeAssignment.RoleId);
                cmd.Parameters.AddWithValue("@explanationId", employeeAssignment.ExplanationId);
                cmd.Parameters.AddWithValue("@companyId", employeeAssignment.CompanyId);
                cmd.Parameters.AddWithValue("@unitPrice", employeeAssignment.UnitPrice);
                cmd.Parameters.AddWithValue("@gradeId", employeeAssignment.GradeId);
                cmd.Parameters.AddWithValue("@updatedBy", employeeAssignment.UpdatedBy);
                cmd.Parameters.AddWithValue("@updatedDate", employeeAssignment.UpdatedDate);

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

        public List<EmployeeAssignmentViewModel>  SearchAssignment(EmployeeAssignment employeeAssignment)
        {
            string query = $@"select ea.id as AssignmentId,ea.SectionId, sec.Name as SectionName,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id";

            List<EmployeeAssignmentViewModel> employeeAssignments = new List<EmployeeAssignmentViewModel>();

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
                            EmployeeAssignmentViewModel employeeAssignmentViewModel = new EmployeeAssignmentViewModel();
                            employeeAssignmentViewModel.Id = Convert.ToInt32(rdr["AssignmentId"]);
                            employeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.RoleName = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();



                            employeeAssignments.Add(employeeAssignmentViewModel);
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