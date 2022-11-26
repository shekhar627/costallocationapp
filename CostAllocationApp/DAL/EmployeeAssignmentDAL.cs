using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class EmployeeAssignmentDAL : DbContext
    {
        public int CreateAssignment(EmployeeAssignment employeeAssignment)
        {
            int result = 0;
            string query = $@"insert into EmployeesAssignments(SectionId,DepartmentId,InChargeId,RoleId,ExplanationId,CompanyId,UnitPrice,GradeId) values(@sectionId,@departmentId,@inChargeId,@roleId,@explanationId,@companyId,@unitPrice,@gradeId,@createdBy,@createdDate)";
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
    }
}