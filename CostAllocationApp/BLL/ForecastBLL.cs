using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class ForecastBLL
    {
        ForecastDAL forecastDAL = null;
        public ForecastBLL()
        {
            forecastDAL = new ForecastDAL();
        }
        public int CreateForecast(Forecast forecast)
        {
            return forecastDAL.CreateForecast(forecast);
        }
        public bool CheckAssignmentId(int assignmentId,int year,int month)
        {
            return forecastDAL.CheckAssignmentId(assignmentId, year, month);
        }
        public int UpdateForecast(Forecast forecast)
        {
            return forecastDAL.UpdateForecast(forecast);
        }

        public int UpdateEmployeesAssignmentFromForecst(string allocationId, string assignmentId)
        {
            return forecastDAL.UpdateEmployeesAssignmentFromForecst(allocationId, assignmentId);
        }
    }
}