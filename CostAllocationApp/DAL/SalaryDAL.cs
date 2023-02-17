using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;
using System.Globalization;
using CostAllocationApp.ViewModels;

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
                            if (string.IsNullOrEmpty(rdr["GradeLowPoints"].ToString()))
                            {
                                salary.SalaryLowPoint = 0;
                            }
                            else
                            {
                                salary.SalaryLowPoint = Convert.ToDecimal(rdr["GradeLowPoints"]);
                                salary.SalaryLowPointWithComma = Convert.ToInt32(rdr["GradeLowPoints"]).ToString("N0");
                            }


                            if (string.IsNullOrEmpty(rdr["GradeHighPoints"].ToString()))
                            {
                                salary.SalaryHighPoint = 0;
                            }
                            else
                            {
                                salary.SalaryHighPoint = Convert.ToDecimal(rdr["GradeHighPoints"]);
                                salary.SalaryHighPointWithComma = Convert.ToInt32(rdr["GradeHighPoints"]).ToString("N0");
                            }

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
        public List<GradeSalaryTypeViewModel> GetAllSalaryTypes()
        {
            List<GradeSalaryTypeViewModel> salaries = new List<GradeSalaryTypeViewModel>();
            string query = "";
            query = query + "select gst.id,gst.GradeId,g.GradeName,gst.GradeLowPoints,gst.GradeHighPoints ";
            query = query + "    ,gst.DepartmentId,d.Name,gst.Year,gst.SalaryTypeId,st.SalaryTypeName ";
            query = query + "from GradeSalarlyTypes gst ";
            query = query + "    join salarytypes st on gst.SalaryTypeId = st.Id ";
            query = query + "    join departments d on gst.DepartmentId = d.Id ";
            query = query + "    join grades g on gst.GradeId = g.Id ";

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
                            GradeSalaryTypeViewModel salary = new GradeSalaryTypeViewModel();
                            salary.Id = Convert.ToInt32(rdr["Id"]);
                            salary.GradeId = Convert.ToInt32(rdr["GradeId"]);
                            salary.GradeName = rdr["GradeName"].ToString();
                            salary.GradeLowPoints = Convert.ToDecimal(rdr["GradeLowPoints"]);
                            salary.GradeHighPoints = Convert.ToDecimal(rdr["GradeHighPoints"]);
                            salary.DepartmentId = Convert.ToInt32(rdr["DepartmentId"]);
                            salary.DepartmentName = rdr["Name"].ToString();
                            salary.SalaryTypeId = Convert.ToInt32(rdr["SalaryTypeId"]);
                            salary.SalaryTypeName = rdr["SalaryTypeName"].ToString();
                            salary.Year = Convert.ToInt32(rdr["Year"]);

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
        public List<GradeSalaryTypeViewModel> GetAllSalaries()
        {
            List<GradeSalaryTypeViewModel> salaries = new List<GradeSalaryTypeViewModel>();
            string query = "";
            query = query + "select id,gradename from grades ";

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
                            GradeSalaryTypeViewModel salary = new GradeSalaryTypeViewModel();
                            salary.GradeId = Convert.ToInt32(rdr["Id"]);
                            salary.GradeName = rdr["gradename"].ToString();

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
        public int GetGradeId(string salaryTypeId)
        {
            List<GradeSalaryTypeViewModel> salaries = new List<GradeSalaryTypeViewModel>();
            GradeSalaryTypeViewModel salary = new GradeSalaryTypeViewModel();
            string query = "";
            query = query + "select id,gradeid from GradeSalarlyTypes where id=" + salaryTypeId;

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
                            salary.Id = Convert.ToInt32(rdr["Id"]);
                            salary.GradeId = Convert.ToInt32(rdr["gradeid"]);

                            salaries.Add(salary);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return salary.GradeId;
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
        public bool CheckGrade(Salary salary)
        {
            string query = $"select * from Grades where GradePoints=N'{salary.SalaryGrade}' or GradeLowPoints={salary.SalaryLowPoint} or GradeHighPoints={salary.SalaryHighPoint}";
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

        public int GetSalaryCountWithEmployeeAsignment(int gradeId)
        {
            string query = "select * from EmployeesAssignments where GradeId=" + gradeId;
            int result = 0;
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
                            result++;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return result;
        }

        public Salary GetSalaryBySalaryId(int salaryId)
        {
            Salary salary = null;
            string query = "select * from Grades where Id = " + salaryId;
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
                        while (rdr.Read())
                        {
                            salary = new Salary();
                            salary.SalaryGrade = rdr["GradePoints"].ToString();
                            salary.Id = Convert.ToInt32(rdr["Id"]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    salary = null;
                }

                return salary;
            }
        }

        public int CreateGradeSalaryType(GradeSalaryType gradeSalaryType)
        {
            int result = 0;
            string query = $@"insert into GradeSalarlyTypes(GradeId,GradeLowPoints,GradeHighPoints,DepartmentId,Year,SalaryTypeId,CreatedBy,CreatedDate) values(@gradeId,@gradeLowPoints,@gradeHighPoints,@departmentId,@year,@salaryTypeId,@createdBy,@createdDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@gradeId", gradeSalaryType.GradeId);
                cmd.Parameters.AddWithValue("@gradeLowPoints", gradeSalaryType.GradeLowPoints);
                cmd.Parameters.AddWithValue("@gradeHighPoints", gradeSalaryType.GradeHighPoints);
                cmd.Parameters.AddWithValue("@departmentId", gradeSalaryType.DepartmentId);
                cmd.Parameters.AddWithValue("@year", gradeSalaryType.Year);
                cmd.Parameters.AddWithValue("@salaryTypeId", gradeSalaryType.SalaryTypeId);
                cmd.Parameters.AddWithValue("@createdBy", gradeSalaryType.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", gradeSalaryType.CreatedDate);
                //cmd.Parameters.AddWithValue("@isActive", salary.IsActive);
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

        public GradeSalaryType GetGradeSalaryType(int departmentId, int salaryTypeId, int year, int gradeId)
        {
            GradeSalaryType gradeSalaryType = new GradeSalaryType(); ;
            string query = $@"select * from GradeSalarlyTypes where GradeId = {gradeId} and DepartmentId={departmentId} and year={year} and SalaryTypeId = {salaryTypeId    }";
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
                        while (rdr.Read())
                        {
                            gradeSalaryType.Id = Convert.ToInt32(rdr["Id"]);
                            gradeSalaryType.GradeId = Convert.ToInt32(rdr["GradeId"]);
                            gradeSalaryType.GradeLowPoints = Convert.ToDouble(rdr["GradeLowPoints"]);
                            gradeSalaryType.GradeHighPoints = Convert.ToDouble(rdr["GradeHighPoints"]);
                            gradeSalaryType.DepartmentId = Convert.ToInt32(rdr["DepartmentId"]);
                            gradeSalaryType.Year = Convert.ToInt32(rdr["Year"]);
                            gradeSalaryType.SalaryTypeId = Convert.ToInt32(rdr["SalaryTypeId"]);

                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return gradeSalaryType;
            }
        }

        public List<GradeSalaryType> GetGradeSalaryTypeByYear_SalaryTypeId_GradeId(int salaryTypeId, int year, int gradeId)
        {
            List<GradeSalaryType> gradeSalaryTypes = new List<GradeSalaryType>();
            string query = $@"select * from GradeSalarlyTypes join Grades on GradeSalarlyTypes.GradeId = Grades.Id where GradeSalarlyTypes.year={year} and GradeSalarlyTypes.SalaryTypeId = {salaryTypeId} and GradeSalarlyTypes.GradeId={gradeId} order by GradeSalarlyTypes.DepartmentId ASC";
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
                        while (rdr.Read())
                        {
                            GradeSalaryType gradeSalaryType = new GradeSalaryType();
                            gradeSalaryType.Id = Convert.ToInt32(rdr["Id"]);
                            gradeSalaryType.GradeId = Convert.ToInt32(rdr["GradeId"]);
                            gradeSalaryType.GradeName = rdr["GradeName"].ToString();
                            gradeSalaryType.GradeLowPoints = Convert.ToDouble(rdr["GradeLowPoints"]);
                            gradeSalaryType.GradeHighPoints = Convert.ToDouble(rdr["GradeHighPoints"]);
                            gradeSalaryType.DepartmentId = Convert.ToInt32(rdr["DepartmentId"]);
                            gradeSalaryType.Year = Convert.ToInt32(rdr["Year"]);
                            gradeSalaryType.SalaryTypeId = Convert.ToInt32(rdr["SalaryTypeId"]);
                            gradeSalaryType.GradeLowWithCommaSeperate = gradeSalaryType.GradeLowPoints.ToString("N0");
                            gradeSalaryType.GradeHighWithCommaSeperate = gradeSalaryType.GradeHighPoints.ToString("N0");

                            gradeSalaryTypes.Add(gradeSalaryType);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return gradeSalaryTypes;
            }
        }
        public List<int> GetSalaryTypeIdByYear(int year)
        {

            List<int> salaryTypeIds = new List<int>();
            string query = $@"select DISTINCT SalaryTypeId from GradeSalarlyTypes WHERE Year={year}";
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
                        while (rdr.Read())
                        {
                            salaryTypeIds.Add(Convert.ToInt32(rdr["SalaryTypeId"]));

                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return salaryTypeIds;
            }
        }
        public GradeSalaryTypeViewModel GetUnitPrice(string gradeId, string departmentId, string year)
        {
            GradeSalaryTypeViewModel _salaryTypeViewModel = new GradeSalaryTypeViewModel();
            string query = "";
            query = "select gst.id,gst.gradeid,gst.departmentid,g.gradename,gst.gradelowpoints ";
            query = query + " from GradeSalarlyTypes gst ";
            query = query + "    join grades g on gst.gradeid = g.id ";
            query = query + "where gst.id =" + gradeId; //+ " and gst.departmentid=" + departmentId + " and gst.year=" + year + " and salarytypeid=2 ";

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
                            _salaryTypeViewModel.Id = Convert.ToInt32(rdr["id"]);
                            _salaryTypeViewModel.GradeId = Convert.ToInt32(rdr["gradeid"]);
                            _salaryTypeViewModel.DepartmentId = Convert.ToInt32(rdr["departmentid"]);
                            _salaryTypeViewModel.GradeName = rdr["gradename"].ToString();
                            _salaryTypeViewModel.GradeLowPoints = Convert.ToInt32(rdr["gradelowpoints"]);

                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return _salaryTypeViewModel;
            }
        }
    }
}