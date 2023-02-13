﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class GradeSalaryType:Common
    {
        public int Id { get; set; }
        public int GradeId { get; set; }
        public double GradeLowPoints { get; set; }
        public double GradeHighPoints { get; set; }
        public int DepartmentId { get; set; }
        public int Year { get; set; }
        public int SalaryTypeId { get; set; }


        // for other usages
        public string GradeName { get; set; }
    }
}