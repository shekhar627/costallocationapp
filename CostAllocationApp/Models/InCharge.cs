using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class InCharge : Common
    {
        public int Id { get; set; }
        public string InChargeName { get; set; }
        public bool IsActive { get; set; }
    }
}