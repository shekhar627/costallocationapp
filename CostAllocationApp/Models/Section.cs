using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class Section : Common
    {
        public int Id { get; set; }
        public string SectionName { get; set; }
        public bool IsActive { get; set; }
    }
}