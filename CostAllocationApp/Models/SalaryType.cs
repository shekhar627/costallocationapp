using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class SalaryType : Common
    {
        public int Id { get; set; }
        public string SalaryTypeName { get; set; }
        public bool IsActive { get; set; }
    }
}