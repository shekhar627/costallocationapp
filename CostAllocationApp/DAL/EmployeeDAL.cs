using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;


namespace CostAllocationApp.DAL
{
    public class EmployeeDAL : DbContext
    {
        public int CreateEmployee(Employee employee)
        {
            int result = 0;
            string query = $@"insert into Employees(FirstName,LastName,Memo,Sex,MobileNo,PresentAddress,PermanentAddress,CreatedBy,CreatedDate,IsActive) values(@firstName,@lastName,@memo,@sex,@mobileNo,@presentAddress,@permanentAddress,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@firstName", employee.FirstName);
                cmd.Parameters.AddWithValue("@lastName", employee.LastName);
                cmd.Parameters.AddWithValue("@memo", employee.Memo);
                cmd.Parameters.AddWithValue("@sex", employee.Sex);
                cmd.Parameters.AddWithValue("@mobileNo", employee.MobileNo);
                cmd.Parameters.AddWithValue("@presentAddress", employee.PresentAddress);
                cmd.Parameters.AddWithValue("@permanentAddress", employee.PermanentAddress);
                cmd.Parameters.AddWithValue("@createdBy", employee.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", employee.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", employee.IsActive);
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
        public List<Employee> GetAllEmployees()
        {
            List<Employee> employees = new List<Employee>();
            string query = "";
            query = "SELECT * ";
            query = query + "FROM Employees ";
            query = query + "WHERE Isactive=1 ";
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
                            Employee employee = new Employee();
                            employee.Id = Convert.ToInt32(rdr["Id"]);
                            employee.FirstName = rdr["FirstName"].ToString();
                            employee.LastName = rdr["LastName"].ToString();
                            employee.Memo = rdr["Memo"].ToString();
                            employee.Sex = rdr["Sex"].ToString();
                            employee.MobileNo = rdr["MobileNo"].ToString();
                            employee.PresentAddress = rdr["PresentAddress"].ToString();
                            employee.PermanentAddress = rdr["PermanentAddress"].ToString();
                            employee.CreatedBy = rdr["CreatedBy"].ToString();
                            employee.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);

                            employees.Add(employee);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return employees;
            }
        }
        public int RemoveEmployee(int employeeIds)
        {
            int result = 0;
            string query = $@"update Departments set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", employeeIds);
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