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
                if (String.IsNullOrEmpty(employeeAssignment.ExplanationId))
                {
                    cmd.Parameters.AddWithValue("@explanationId", DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@explanationId", employeeAssignment.ExplanationId);
                }

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
                if (String.IsNullOrEmpty(employeeAssignment.ExplanationId))
                {
                    cmd.Parameters.AddWithValue("@explanationId", DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@explanationId", employeeAssignment.ExplanationId);
                }

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

        public List<EmployeeAssignmentViewModel> SearchAssignment(EmployeeAssignment employeeAssignment)
        {
            string where = "";
            if (employeeAssignment.SectionId >0)
            {
                where += $" ea.SectionId={employeeAssignment.SectionId} and ";
            }
            if (employeeAssignment.DepartmentId >0)
            {
                where += $" ea.DepartmentId={employeeAssignment.DepartmentId} and ";
            }
            if (employeeAssignment.InchargeId >0)
            {
                where += $" ea.InChargeId={employeeAssignment.InchargeId} and ";
            }
            if (employeeAssignment.RoleId > 0)
            {
                where += $" ea.RoleId={employeeAssignment.RoleId} and ";
            }
            //if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
            //{
            //    where += $" ea.ExplanationId={employeeAssignment.ExplanationId} and ";
            //}
            if (employeeAssignment.CompanyId >0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }

            where += " 1=1 ";
            string query = $@"select ea.id as AssignmentId,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode, ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
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
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
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

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks,gd.GradePoints,ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice, ea.GradeId 
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
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
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
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
            if (employeeAssignment.DepartmentId >0)
            {
                where += $" ea.DepartmentId={employeeAssignment.DepartmentId} and ";
            }
            if (employeeAssignment.InchargeId >0)
            {
                where += $" ea.InChargeId={employeeAssignment.InchargeId} and ";
            }
            if (employeeAssignment.RoleId >0)
            {
                where += $" ea.RoleId={employeeAssignment.RoleId} and ";
            }
            //if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
            //{
            //    where += $" ea.ExplanationId={employeeAssignment.ExplanationId} and ";
            //}
            if (employeeAssignment.CompanyId >0)
            {
                where += $" ea.CompanyId={employeeAssignment.CompanyId} and ";
            }
            
            if (employeeAssignment.IsActive == "0" || employeeAssignment.IsActive == "1")
            {
                where += $" ea.IsActive={employeeAssignment.IsActive} and ";
            }
            else
            {
                where += $" ea.IsActive=1 and ";
            }

            where += " 1=1 ";

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode, ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id 
                            join Grades gd on ea.GradeId = gd.Id
                            where {where}
                            order by ea.EmployeeName asc";
            //ORDER BY ea.EmployeeName asc, ea.Id";


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
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToDecimal(rdr["UnitPrice"]).ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));

                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            //employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeName;
                            employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            employeeAssignmentViewModel.Remarks = rdr["Remarks"] is DBNull ? "" :  rdr["Remarks"].ToString();
                            //if (!string.IsNullOrEmpty(rdr["SubCode"].ToString()))
                            //{
                            //    employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            //    employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " " + employeeAssignmentViewModel.SubCode;
                            //}
                            //else
                            //{
                            //    employeeAssignmentViewModel.SubCode = 0;                                
                            //}

                            //if (!string.IsNullOrEmpty(rdr["Remarks"].ToString()))
                            //{
                            //    employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            //    employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " (" + employeeAssignmentViewModel.Remarks + ")";
                            //}
                            //else
                            //{
                            //    employeeAssignmentViewModel.Remarks = "";
                            //}


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

        public List<ForecastAssignmentViewModel> GetEmployeesForecastBySearchFilter(EmployeeAssignmentForecast employeeAssignment)
        {

            string where = "";
            if (!string.IsNullOrEmpty(employeeAssignment.EmployeeName))
            {
                where += $" ea.EmployeeName like N'%{employeeAssignment.EmployeeName}%' and ";
            }
            if (!string.IsNullOrEmpty(employeeAssignment.SectionId))
            {
                string tempSectionIds = "";
                if (employeeAssignment.SectionId.IndexOf("##") >= 0)
                {
                    string[] arrSectionIds = employeeAssignment.SectionId.Split(new[] { "##" }, StringSplitOptions.None);
                    
                    for (int i = 0; i < arrSectionIds.Length; i++)
                    {
                        if(tempSectionIds == "")
                        {
                            tempSectionIds = arrSectionIds[i];
                        }
                        else
                        {
                            tempSectionIds = tempSectionIds +","+arrSectionIds[i];
                        }
                    }
                }
                else
                {
                    tempSectionIds = employeeAssignment.SectionId;
                }
                where += $" ea.SectionId In ({tempSectionIds}) and ";
            }
            if (!string.IsNullOrEmpty(employeeAssignment.DepartmentId))
            {
                string tempDepartmentIds = "";
                if (employeeAssignment.DepartmentId.IndexOf("##") >= 0)
                {
                    string[] arrDepartmentIds = employeeAssignment.DepartmentId.Split(new[] { "##" }, StringSplitOptions.None);

                    for (int i = 0; i < arrDepartmentIds.Length; i++)
                    {
                        if (tempDepartmentIds == "")
                        {
                            tempDepartmentIds = arrDepartmentIds[i];
                        }
                        else
                        {
                            tempDepartmentIds = tempDepartmentIds + "," + arrDepartmentIds[i];
                        }
                    }
                }
                else
                {
                    tempDepartmentIds = employeeAssignment.DepartmentId;
                }
                where += $" ea.DepartmentId In ({tempDepartmentIds}) and ";
            }

            if (!string.IsNullOrEmpty(employeeAssignment.InchargeId))
            {
                string tempInchargeIdIds = "";
                if (employeeAssignment.InchargeId.IndexOf("##") >= 0)
                {
                    string[] arrInchargeIds = employeeAssignment.InchargeId.Split(new[] { "##" }, StringSplitOptions.None);

                    for (int i = 0; i < arrInchargeIds.Length; i++)
                    {
                        if (tempInchargeIdIds == "")
                        {
                            tempInchargeIdIds = arrInchargeIds[i];
                        }
                        else
                        {
                            tempInchargeIdIds = tempInchargeIdIds + "," + arrInchargeIds[i];
                        }
                    }
                }
                else
                {
                    tempInchargeIdIds = employeeAssignment.InchargeId;
                }
                where += $" ea.InChargeId In ({tempInchargeIdIds}) and ";
            }

            if (!string.IsNullOrEmpty(employeeAssignment.RoleId))
            {
                string tempRoleIds = "";
                if (employeeAssignment.RoleId.IndexOf("##") >= 0)
                {
                    string[] arrRoleIdss = employeeAssignment.RoleId.Split(new[] { "##" }, StringSplitOptions.None);

                    for (int i = 0; i < arrRoleIdss.Length; i++)
                    {
                        if (tempRoleIds == "")
                        {
                            tempRoleIds = arrRoleIdss[i];
                        }
                        else
                        {
                            tempRoleIds = tempRoleIds + "," + arrRoleIdss[i];
                        }
                    }
                }
                else
                {
                    tempRoleIds = employeeAssignment.RoleId;
                }
                where += $" ea.RoleId In ({tempRoleIds}) and ";
            }

            if (!String.IsNullOrEmpty(employeeAssignment.ExplanationId))
            {
                string tempExplanationIds = "";
                if (employeeAssignment.ExplanationId.IndexOf("##") >= 0)
                {
                    string[] arrtempExplanationIds = employeeAssignment.ExplanationId.Split(new[] { "##" }, StringSplitOptions.None);

                    for (int i = 0; i < arrtempExplanationIds.Length; i++)
                    {
                        if (tempExplanationIds == "")
                        {
                            tempExplanationIds = arrtempExplanationIds[i];
                        }
                        else
                        {
                            tempExplanationIds = tempExplanationIds + "," + arrtempExplanationIds[i];
                        }
                    }
                }
                else
                {
                    tempExplanationIds = employeeAssignment.ExplanationId;
                }
                where += $" ea.ExplanationId IN ({tempExplanationIds}) and ";
            }
            if (!string.IsNullOrEmpty(employeeAssignment.CompanyId))
            {
                string tempCompanyIds = "";
                if (employeeAssignment.CompanyId.IndexOf("##") >= 0)
                {
                    string[] arrCompanyIds = employeeAssignment.CompanyId.Split(new[] { "##" }, StringSplitOptions.None);

                    for (int i = 0; i < arrCompanyIds.Length; i++)
                    {
                        if (tempCompanyIds == "")
                        {
                            tempCompanyIds = arrCompanyIds[i];
                        }
                        else
                        {
                            tempCompanyIds = tempCompanyIds + "," + arrCompanyIds[i];
                        }
                    }
                }
                else
                {
                    tempCompanyIds = employeeAssignment.CompanyId;
                }
                where += $" ea.CompanyId In ({tempCompanyIds}) and ";
            }

            if (employeeAssignment.IsActive == "0" || employeeAssignment.IsActive == "1")
            {
                where += $" ea.IsActive={employeeAssignment.IsActive} and ";
            }
            else
            {
                where += $" ea.IsActive=1 and ";
            }

            where += " 1=1 ";

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode, ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            join Roles rl on ea.RoleId = rl.Id
                            join InCharges inc on ea.InChargeId = inc.Id 
                            join Grades gd on ea.GradeId = gd.Id
                            where {where}
                            order by ea.EmployeeName asc";


            List<ForecastAssignmentViewModel> forecastEmployeeAssignments = new List<ForecastAssignmentViewModel>();
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
                            ForecastAssignmentViewModel forecastEmployeeAssignmentViewModel = new ForecastAssignmentViewModel();
                            forecastEmployeeAssignmentViewModel.Id = Convert.ToInt32(rdr["AssignmentId"]);
                            forecastEmployeeAssignmentViewModel.EmployeeName = rdr["EmployeeName"].ToString();
                            forecastEmployeeAssignmentViewModel.SectionId = rdr["SectionId"].ToString();
                            forecastEmployeeAssignmentViewModel.SectionName = rdr["SectionName"].ToString();
                            forecastEmployeeAssignmentViewModel.DepartmentId = rdr["DepartmentId"].ToString();
                            forecastEmployeeAssignmentViewModel.DepartmentName = rdr["DepartmentName"].ToString();
                            forecastEmployeeAssignmentViewModel.InchargeId = rdr["InchargeId"].ToString();
                            forecastEmployeeAssignmentViewModel.InchargeName = rdr["InchargeName"].ToString();
                            forecastEmployeeAssignmentViewModel.RoleId = rdr["RoleId"].ToString();
                            forecastEmployeeAssignmentViewModel.RoleName = rdr["RoleName"].ToString();
                            forecastEmployeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            forecastEmployeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            forecastEmployeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            forecastEmployeeAssignmentViewModel.UnitPrice = Convert.ToDecimal(rdr["UnitPrice"]).ToString();
                            forecastEmployeeAssignmentViewModel.UnitPrice = Convert.ToInt32(forecastEmployeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));

                            forecastEmployeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            forecastEmployeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            forecastEmployeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            forecastEmployeeAssignmentViewModel.Remarks = rdr["Remarks"]  is DBNull ? "":  rdr["Remarks"].ToString() ;

                            //if (!string.IsNullOrEmpty(rdr["SubCode"].ToString()))
                            //{
                            //    forecastEmployeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            //    forecastEmployeeAssignmentViewModel.EmployeeNameWithCodeRemarks = forecastEmployeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " " + forecastEmployeeAssignmentViewModel.SubCode;
                            //}
                            //else
                            //{
                            //    forecastEmployeeAssignmentViewModel.SubCode = 0;
                            //}

                            //if (!string.IsNullOrEmpty(rdr["Remarks"].ToString()))
                            //{
                            //    forecastEmployeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            //    forecastEmployeeAssignmentViewModel.EmployeeNameWithCodeRemarks = forecastEmployeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " (" + forecastEmployeeAssignmentViewModel.Remarks + ")";
                            //}
                            //else
                            //{
                            //    forecastEmployeeAssignmentViewModel.Remarks = "";
                            //}

                            //if (!string.IsNullOrEmpty(rdr["Remarks"].ToString()))
                            //{
                            //    forecastEmployeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            //}
                            //else
                            //{
                            //    forecastEmployeeAssignmentViewModel.Remarks = "";
                            //}
                            //if (!string.IsNullOrEmpty(rdr["SubCode"].ToString()))
                            //{
                            //    forecastEmployeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            //}
                            //else
                            //{
                            //    forecastEmployeeAssignmentViewModel.SubCode = 0;
                            //}

                            //HttpContext.Current.Response.Write("employeeAssignmentViewModel.UnitPrice: " + employeeAssignmentViewModel.UnitPrice);
                            //HttpContext.Current.Response.End();

                            forecastEmployeeAssignments.Add(forecastEmployeeAssignmentViewModel);
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return forecastEmployeeAssignments;
        }


        public List<EmployeeAssignmentViewModel> GetEmployeesBySearchFilterForMultipleSearch(EmployeeAssignmentDTO employeeAssignment)
        {

            string where = "";
            if (!string.IsNullOrEmpty(employeeAssignment.EmployeeName))
            {
                where += $" ea.EmployeeName like N'%{employeeAssignment.EmployeeName.Trim()}%' and ";
            }
            if (employeeAssignment.Sections != null)
            {
                if (employeeAssignment.Sections.Length > 0 && employeeAssignment.Sections.ToString() != "all")
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
            if (employeeAssignment.Departments != null && employeeAssignment.Departments.ToString() != "all")
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

            if (employeeAssignment.Incharges != null && employeeAssignment.Incharges.ToString() != "all")
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
            if (employeeAssignment.Roles != null && employeeAssignment.Roles.ToString() != "all")
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

            //if (employeeAssignment.Explanations != null)
            //{
            //    if (employeeAssignment.Explanations.Length > 0)
            //    {
            //        string ids = "";
            //        foreach (var item in employeeAssignment.Explanations)
            //        {
            //            ids += $"{item},";
            //        }
            //        ids = ids.TrimEnd(',');

            //        where += $" ea.ExplanationId in ({ids}) and ";
            //    }

            //}
            if (employeeAssignment.Companies != null && employeeAssignment.Companies.ToString() != "all")
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
            else
            {
                where += $" ea.IsActive=1 and ";
            }

            where += " 1=1 ";
            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode,ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
                            
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
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToDecimal(rdr["UnitPrice"]).ToString("N2");
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeName;

                            if (!string.IsNullOrEmpty(rdr["SubCode"].ToString()))
                            {
                                employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                                employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " " + employeeAssignmentViewModel.SubCode;
                            }
                            else
                            {
                                employeeAssignmentViewModel.SubCode = 0;
                            }

                            if (!string.IsNullOrEmpty(rdr["Remarks"].ToString()))
                            {
                                employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                                employeeAssignmentViewModel.EmployeeNameWithCodeRemarks = employeeAssignmentViewModel.EmployeeNameWithCodeRemarks + " (" + employeeAssignmentViewModel.Remarks + ")";
                            }
                            else
                            {
                                employeeAssignmentViewModel.Remarks = "";
                            }




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

            string query = $@"select ea.id as AssignmentId,ea.EmployeeName,ea.SectionId, sec.Name as SectionName, ea.Remarks, ea.SubCode,ea.ExplanationId,
                            ea.DepartmentId, dep.Name as DepartmentName,ea.InChargeId, inc.Name as InchargeName,ea.RoleId,rl.Name as RoleName,ea.CompanyId, com.Name as CompanyName, ea.UnitPrice
                            ,gd.GradePoints,ea.IsActive
                            from EmployeesAssignments ea join Sections sec on ea.SectionId = sec.Id
                            join Departments dep on ea.DepartmentId = dep.Id
                            join Companies com on ea.CompanyId = com.Id
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
                            employeeAssignmentViewModel.ExplanationId = rdr["ExplanationId"] is DBNull ? "" : rdr["ExplanationId"].ToString();
                            //employeeAssignmentViewModel.ExplanationName = rdr["ExplanationName"] is DBNull ? "" : rdr["ExplanationName"].ToString();
                            employeeAssignmentViewModel.CompanyId = rdr["CompanyId"].ToString();
                            employeeAssignmentViewModel.CompanyName = rdr["CompanyName"].ToString();
                            employeeAssignmentViewModel.UnitPrice = rdr["UnitPrice"].ToString();
                            employeeAssignmentViewModel.UnitPrice = Convert.ToInt32(employeeAssignmentViewModel.UnitPrice).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
                            employeeAssignmentViewModel.UnitPriceWithoutComma = Convert.ToInt32(rdr["UnitPrice"].ToString());
                            employeeAssignmentViewModel.GradePoint = rdr["GradePoints"].ToString();
                            employeeAssignmentViewModel.IsActive = Convert.ToBoolean(rdr["IsActive"]);
                            if (!string.IsNullOrEmpty(rdr["Remarks"].ToString()))
                            {
                                employeeAssignmentViewModel.Remarks = rdr["Remarks"].ToString();
                            }
                            else
                            {
                                employeeAssignmentViewModel.Remarks = "";
                            }
                            if (!string.IsNullOrEmpty(rdr["SubCode"].ToString()))
                            {
                                employeeAssignmentViewModel.SubCode = Convert.ToInt32(rdr["SubCode"]);
                            }
                            else
                            {
                                employeeAssignmentViewModel.SubCode = 0;
                            }



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

        public List<ForecastDto> GetForecastsByAssignmentId(int assignmentId)
        {
            List<ForecastDto> forecasts = new List<ForecastDto>();
            string query = "select * from Costs where EmployeeAssignmentsId=" + assignmentId;
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
                            ForecastDto forecast = new ForecastDto();
                            forecast.ForecastId = Convert.ToInt32(rdr["Id"]);
                            forecast.Year = Convert.ToInt32(rdr["Year"]);
                            forecast.Month = Convert.ToInt32(rdr["MonthId"]);
                            forecast.Points = Convert.ToDecimal(rdr["Points"]);
                            //forecast.Total = Convert.ToDecimal(rdr["Total"]);
                            forecast.Total = rdr["Total"].ToString();
                            forecast.Total = Convert.ToDecimal(forecast.Total).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
                            if (String.IsNullOrEmpty(forecast.Total))
                            {
                                forecast.Total = "0";
                            }
                            //if (!string.IsNullOrEmpty(rdr["Total"].ToString())) {
                            //    forecast.Total = rdr["Total"].ToString();
                            //    //forecast.Total = (Convert.ToDecimal(forecast.Total)).ToString("N", new CultureInfo("en-US"));
                            //    forecast.Total = Convert.ToDecimal(forecast.Total).ToString("#,#.##", CultureInfo.CreateSpecificCulture("hi-IN"));
                            //}
                            //else
                            //{
                            //    forecast.Total = "0";
                            //}

                            forecasts.Add(forecast);
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return forecasts;
        }

        public bool CheckEmployeeName(string employeeName)
        {
            string query = "select * from EmployeesAssignments where EmployeeName=N'"+employeeName+"'";
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

    }
}