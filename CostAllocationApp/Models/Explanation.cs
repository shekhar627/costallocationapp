using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class Explanation : Common    {
        public int Id { get; set; }
        public string ExplanationName { get; set; }
        public bool IsActive { get; set; }
    }
}