﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class UnitPriceType : Common
    {
        public int Id { get; set; }
        public string UnitPriceTypeName { get; set; }
        public bool IsActive { get; set; }
    }
}