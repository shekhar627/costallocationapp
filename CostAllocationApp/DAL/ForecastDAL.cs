using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;
using System.Data.SqlClient;


namespace CostAllocationApp.DAL
{
    public class ForecastDAL : DbContext
    {
        public int CreateForecast(Forecast forecast)
        {
            int result = 0;
            string query = $@"insert into Costs(Year,MonthId,Points,Total,EmployeeAssignmentsId,CreatedBy,CreatedDate) values(@year,@monthId,@points,@total,@employeeAssignmentsId,@createdBy,@createdDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@year", forecast.Year);
                cmd.Parameters.AddWithValue("@monthId", forecast.Month);
                cmd.Parameters.AddWithValue("@points", forecast.Points);
                cmd.Parameters.AddWithValue("@total", forecast.Total);
                cmd.Parameters.AddWithValue("@employeeAssignmentsId", forecast.EmployeeAssignmentId);
                cmd.Parameters.AddWithValue("@createdBy", forecast.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", forecast.CreatedDate);
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

        public int UpdateForecast(Forecast forecast)
        {
            int result = 0;
            string query = $@"update costs set Points = @points, Total= @total, UpdatedBy=@updatedBy, UpdatedDate=@updatedDate where Year=@year and EmployeeAssignmentsId=@employeeAssignmentsId and MonthId=@monthId";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                
                
                cmd.Parameters.AddWithValue("@points", forecast.Points);
                cmd.Parameters.AddWithValue("@total", forecast.Total);
                cmd.Parameters.AddWithValue("@year", forecast.Year);
                cmd.Parameters.AddWithValue("@employeeAssignmentsId", forecast.EmployeeAssignmentId);
                cmd.Parameters.AddWithValue("@monthId", forecast.Month);
                
                cmd.Parameters.AddWithValue("@updatedBy", forecast.UpdatedBy);
                cmd.Parameters.AddWithValue("@updatedDate", forecast.UpdatedDate);
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


        public bool CheckAssignmentId(int assignmentId,int year,int month)
        {
            string query = "select * from costs where EmployeeAssignmentsId="+assignmentId+" and year = "+year+" and monthid="+month;
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

        public int UpdateEmployeesAssignmentFromForecst(string allocationId,string assignmentId)
        {
            int result = 0;
            string query = $@"update EmployeesAssignments set ExplanationId = @allocationId where Id = @assignmentId";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                if (String.IsNullOrEmpty(allocationId))
                {
                    cmd.Parameters.AddWithValue("@allocationId", DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@allocationId", allocationId);
                }
                
                cmd.Parameters.AddWithValue("@assignmentId", assignmentId);
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