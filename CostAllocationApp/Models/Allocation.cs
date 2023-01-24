using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class Allocation : Common
    {
        public int Id { get; set; }
        public string AllocationName { get; set; }
        public bool IsActive { get; set; }
    }
}