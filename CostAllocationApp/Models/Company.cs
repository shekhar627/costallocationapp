using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class Company: Common
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public bool IsActive { get; set; }
    }
}