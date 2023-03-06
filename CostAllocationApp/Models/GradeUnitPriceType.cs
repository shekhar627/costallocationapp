using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class GradeUnitPriceType:Common
    {
        public int Id { get; set; }
        public int GradeId { get; set; }
        public double GradeLowPoints { get; set; }
        public double GradeHighPoints { get; set; }
        public int DepartmentId { get; set; }
        public int Year { get; set; }
        public int UnitPriceTypeId { get; set; }


        // for other usages
        public string GradeName { get; set; }
        public string GradeLowWithCommaSeperate { get; set; }
        public string GradeHighWithCommaSeperate { get; set; }
    }
}