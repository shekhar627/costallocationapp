using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Dtos
{
    public class SendToForecaseApiDto
    {
        public string Data { get; set; }
        public int Year { get; set; }
        public int AssignmentId { get; set; }
        public int? AllocationId { get; set; }
    }
}