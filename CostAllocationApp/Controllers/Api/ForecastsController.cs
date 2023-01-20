using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;


namespace CostAllocationApp.Controllers.Api
{
    public class ForecastsController : ApiController
    {
        ForecastBLL forecastBLL = null;
        public ForecastsController()
        {
            forecastBLL = new ForecastBLL();
        }
        [HttpGet]
        public IHttpActionResult CreateForecast(string data,string year,string assignmentId)
        {
            string[] monthData = data.Split(',');
            int tempYear = 0;
            int tempAssignmentId = 0;
            if (!Int32.TryParse(year,out tempYear))
            {
                return BadRequest("Something Went Wrong");
            }
            if (!Int32.TryParse(assignmentId, out tempAssignmentId))
            {
                return BadRequest("Something Went Wrong");
            }

            foreach (var item in monthData)
            {
                string[] temp = item.Split('_');
                Forecast forecast = new Forecast();
                decimal tempPoint = 0;
                if (Decimal.TryParse(temp[1],out tempPoint))
                {
                    if (tempPoint>1 || tempPoint<0)
                    {
                        forecast.Points = 0;
                        forecast.Total = 0;

                    }
                    else
                    {
                        forecast.Points = tempPoint;
                        forecast.Total = Convert.ToDecimal(temp[2]);
                    }
                }
                else
                {
                    forecast.Points = tempPoint;
                    forecast.Total = 0;
                }

                forecast.Month = Convert.ToInt32(temp[0]);
                forecast.Year = Convert.ToInt32(year);
                forecast.EmployeeAssignmentId = Convert.ToInt32(assignmentId);
                forecast.CreatedBy = "";
                forecast.CreatedDate = DateTime.Now;
                forecast.UpdatedBy = "";
                forecast.UpdatedDate = DateTime.Now;

                var result = forecastBLL.CheckAssignmentId(int.Parse(assignmentId),int.Parse(year), Convert.ToInt32(temp[0]));
                if (result==true)
                {
                    int resultEdit = forecastBLL.UpdateForecast(forecast);
                }
                else
                {
                    int resultSave = forecastBLL.CreateForecast(forecast);
                }

                
            }
            return Ok(true);

        }

        [HttpPost]
        public IHttpActionResult CreateForecastRecord(string data, string year, string assignmentId)
        {
            string[] monthData = data.Split(',');
            int tempYear = 0;
            int tempAssignmentId = 0;
            if (!Int32.TryParse(year, out tempYear))
            {
                return BadRequest("Something Went Wrong");
            }
            if (!Int32.TryParse(assignmentId, out tempAssignmentId))
            {
                return BadRequest("Something Went Wrong");
            }

            foreach (var item in monthData)
            {
                string[] temp = item.Split('_');
                Forecast forecast = new Forecast();
                decimal tempPoint = 0;
                if (Decimal.TryParse(temp[1], out tempPoint))
                {
                    if (tempPoint > 1 || tempPoint < 0)
                    {
                        forecast.Points = 0;
                        forecast.Total = 0;

                    }
                    else
                    {
                        forecast.Points = tempPoint;
                        forecast.Total = Convert.ToDecimal(temp[2]);
                    }
                }
                else
                {
                    forecast.Points = tempPoint;
                    forecast.Total = 0;
                }

                forecast.Month = Convert.ToInt32(temp[0]);
                forecast.Year = Convert.ToInt32(year);
                forecast.EmployeeAssignmentId = Convert.ToInt32(assignmentId);
                forecast.CreatedBy = "";
                forecast.CreatedDate = DateTime.Now;
                forecast.UpdatedBy = "";
                forecast.UpdatedDate = DateTime.Now;

                var result = forecastBLL.CheckAssignmentId(int.Parse(assignmentId), int.Parse(year), Convert.ToInt32(temp[0]));
                if (result == true)
                {
                    int resultEdit = forecastBLL.UpdateForecast(forecast);
                }
                else
                {
                    int resultSave = forecastBLL.CreateForecast(forecast);
                }


            }
            return Ok(true);

        }
    }
}
