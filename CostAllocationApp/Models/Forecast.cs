using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class Forecast : Common
    {
        public int Id { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public decimal Points { get; set; }
        public decimal Total { get; set; }
        public int EmployeeAssignmentId { get; set; }

    }
}