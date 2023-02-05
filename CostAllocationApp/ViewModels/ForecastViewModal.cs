using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;

namespace CostAllocationApp.ViewModels
{
    public class ForecastViewModal:BaseViewModel
    {
        //public List<Section> _sections { get; set; }
        public Department Department { get; set; }
    }
}