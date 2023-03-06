using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Dtos
{
    public class ForecastDto
    {
        public int ForecastId { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public decimal Points { get; set; }
        //public decimal Total { get; set; }
        public string Total { get; set; }
        public decimal OverTime { get; set; }
    }
}