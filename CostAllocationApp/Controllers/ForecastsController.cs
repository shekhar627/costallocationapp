using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CostAllocationApp.Controllers
{
    public class ForecastsController : Controller
    {
        // GET: Forecasts
        public ActionResult CreateForecast()
        {
            return View();
        }
    }
}