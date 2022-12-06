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
    public class EmployeeAssignmentDAL : DbContext
    {
        public int CreateAssignment(EmployeeAssignment employeeAssignment)
        {
            int result = 0;
            string query = $@"insert into EmployeesAssignments(EmployeeName,SectionId,DepartmentId,InChargeId,RoleId,ExplanationId,CompanyId,UnitPrice,GradeId,CreatedBy,CreatedDate,IsActive,Remarks,SubCode) values(@employeeName,@sectionId,@departmentId,@inChargeId,@roleId,@explanationId,@companyId,@unitPrice,@gradeId,@createdBy,@createdDate,@isActive,@remarks,@subCode)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@employeeName", employeeAssignment.EmployeeName);
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
                cmd.Parameters.AddWithValue("@isActive", employeeAssignment.IsActive);
                cmd.Parameters.AddWithValue("@remarks", employeeAssignment.Remarks);
                cmd.Parameters.AddWithValue("@subCode", employeeAssignment.SubCode);

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
            string query = $@"update EmployeesAssignments set  SectionId=@sectionId,DepartmentId=@departmentId,InChargeId=@inChargeId,RoleId=@roleId,ExplanationId=@explanationId,CompanyId=@companyId,UnitPrice=@unitPrice,GradeId=@gradeId,UpdatedBy=@updatedBy,UpdatedDate=@updatedDate, Remarks=@remarks where Id=@id";
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
                cmd.Parameters.AddWithValue("@id", employeeAssignment.Id);
                cmd.Parameters.AddWithValue("@remarks", employeeAssignment.Remarks);

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
            string where="";
            if (employeeAssignment.SectionId>0)
            {
                where += $" ea.SectionId={employeeAssignment.SectionId} and ";
            }
            if (employeeAssignment.DepartmentId > 0)
            {
                where += $" ea.DepartmentId={employeeAssignment.DepartmentId} and ";
            }
            if (employeeAssignment.InchargeId > 0)
            {
                where += $" ea.InChargeId={employeeAssignment.InchargeId} and ";
            }
            if (employeeAssignment.RoleId > 0)
            {
                where += $" ea.RoleId={employeeAssignment.RoleId} and ";
            }
            if (employeeAssignment.ExplanationId > 0)
            {
                where += $" ea.ExplanationId={employeeAssignment.ExplanationId} and ";
            }
            if (employeeAssignment.CompanyId > 0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }
            if (employeeAssignment.CompanyId > 0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }

            where += " 1=1 ";
            string query = $@"select ea.id as AssignmentId,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id where {where}";

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
                            employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);



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

        public EmployeeAssignmentViewModel GetAssignmentById(int assignmentId)
        {

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks,gd.GradePoints,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice, ea.GradeId 
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join Grades gd on ea.GradeId = gd.Id 
                            join InCharges inc on ea.InChargeId = inc.Id where ea.Id={assignmentId}";

            EmployeeAssignmentViewModel employeeAssignmentViewModel = new EmployeeAssignmentViewModel();

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
                            employeeAssignmentViewModel.EmployeeName = rdr["EmployeeName"].ToString();
                            employeeAssignmentViewModel.Id = Convert.ToInt32(rdr["AssignmentId"]);
                            employeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.GradeId = rdr["GradeId"].ToString();
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();

                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return employeeAssignmentViewModel;
        }

        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilter(EmployeeAssignment employeeAssignment)
        {

            string where = "";
            if (!string.IsNullOrEmpty(employeeAssignment.EmployeeName))
            {
                where += $" ea.EmployeeName like N'%{employeeAssignment.EmployeeName}%' and ";
            }
            if (employeeAssignment.SectionId > 0)
            {
                where += $" ea.SectionId={employeeAssignment.SectionId} and ";
            }
            if (employeeAssignment.DepartmentId > 0)
            {
                where += $" ea.DepartmentId={employeeAssignment.DepartmentId} and ";
            }
            if (employeeAssignment.InchargeId > 0)
            {
                where += $" ea.InChargeId={employeeAssignment.InchargeId} and ";
            }
            if (employeeAssignment.RoleId > 0)
            {
                where += $" ea.RoleId={employeeAssignment.RoleId} and ";
            }
            if (employeeAssignment.ExplanationId > 0)
            {
                where += $" ea.ExplanationId={employeeAssignment.ExplanationId} and ";
            }
            if (employeeAssignment.CompanyId > 0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }
            if (employeeAssignment.CompanyId > 0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }
            if (employeeAssignment.IsActive =="0" || employeeAssignment.IsActive == "1")
            {
                where += $" ea.IsActive={employeeAssignment.IsActive} and ";
            }
            else
            {
                where += $" ea.IsActive=1 and ";
            }

            where += " 1=1 ";
            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id 
                            join Grades gd on ea.GradeId = gd.Id
                            where {where}";

            List<EmployeeAssignmentViewModel> employeeAssignments = new List<EmployeeAssignmentViewModel>();
            //HttpContext.Current.Response.Write("query: " + query + "<br>");
            //HttpContext.Current.Response.End();

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
                            employeeAssignmentViewModel.EmployeeName = rdr["EmployeeName"].ToString();
                            employeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToDecimal(rdr["UnitPrice"]).ToString("N2");
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);

                            //HttpContext.Current.Response.Write("employeeAssignmentViewModel.UnitPrice: " + employeeAssignmentViewModel.UnitPrice);
                            //HttpContext.Current.Response.End();

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

        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilterForMultipleSearch(EmployeeAssignmentDTO employeeAssignment)
        {

            string where = "";
            if (!string.IsNullOrEmpty(employeeAssignment.EmployeeName))
            {
                where += $" ea.EmployeeName like N'%{employeeAssignment.EmployeeName.Trim()}%' and ";
            }
            if (employeeAssignment.Sections!=null)
            {
                if (employeeAssignment.Sections.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Sections)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.SectionId in ({ids}) and ";
                }
              
            }
            if (employeeAssignment.Departments != null)
            {
                if (employeeAssignment.Departments.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Departments)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.DepartmentId in ({ids}) and ";
                }

            }

            if (employeeAssignment.Incharges != null)
            {
                if (employeeAssignment.Incharges.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Incharges)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.InChargeId in ({ids}) and ";
                }

            }
            if (employeeAssignment.Roles != null)
            {
                if (employeeAssignment.Roles.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Roles)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.RoleId in ({ids}) and ";
                }

            }

            if (employeeAssignment.Explanations != null)
            {
                if (employeeAssignment.Explanations.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Explanations)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.ExplanationId in ({ids}) and ";
                }

            }
            if (employeeAssignment.Companies != null)
            {
                if (employeeAssignment.Companies.Length > 0)
                {
                    string ids = "";
                    foreach (var item in employeeAssignment.Companies)
                    {
                        ids += $"{item},";
                    }
                    ids = ids.TrimEnd(',');

                    where += $" ea.CompanyId in ({ids}) and ";
                }

            }
            //if (employeeAssignment.DepartmentId > 0)
            //{
            //    where += $" ea.DepartmentId={employeeAssignment.DepartmentId} and ";
            //}
            //if (employeeAssignment.InchargeId > 0)
            //{
            //    where += $" ea.InChargeId={employeeAssignment.InchargeId} and ";
            //}
            //if (employeeAssignment.RoleId > 0)
            //{
            //    where += $" ea.RoleId={employeeAssignment.RoleId} and ";
            //}
            //if (employeeAssignment.ExplanationId > 0)
            //{
            //    where += $" ea.ExplanationId={employeeAssignment.ExplanationId} and ";
            //}
            //if (employeeAssignment.CompanyId > 0)
            //{
            //    where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            //}
            //if (employeeAssignment.CompanyId > 0)
            //{
            //    where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            //}
            //if (employeeAssignment.IsActive == "0" || employeeAssignment.IsActive == "1")
            //{
            //    where += $" ea.IsActive={employeeAssignment.IsActive} and ";
            //}
            else
            {
                where += $" ea.IsActive=1 and ";
            }

            where += " 1=1 ";
            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id 
                            join Grades gd on ea.GradeId = gd.Id
                            where {where}";

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
                            employeeAssignmentViewModel.EmployeeName = rdr["EmployeeName"].ToString();
                            employeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToDecimal(rdr["UnitPrice"]).ToString("N2");
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);



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


        public List<EmployeeAssignmentViewModel> GetEmployeesByName(string employeeName)
        {

            string where = $"ea.EmployeeName = N'{employeeName}'";

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.ExplanationId, ex.Name as ExplanationName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Explanations ex on ea.ExplanationId = ex.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id 
                            join Grades gd on ea.GradeId = gd.Id
                            where {where} order by ea.SubCode asc";

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
                            employeeAssignmentViewModel.EmployeeName = rdr["EmployeeName"].ToString();
                            employeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            employeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            employeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            employeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            employeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            employeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            employeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            employeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"].ToString();
                            employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);


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

        public int RemoveAssignment(int rowId)
        {
            int result = 0;
            string query = $@"update EmployeesAssignments set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", rowId);
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