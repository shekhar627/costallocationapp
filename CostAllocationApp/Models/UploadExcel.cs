using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class UploadExcel
    {
        public string EmployeeName { get; set; }
        public int? SectionId { get; set; }
        public int? DepartmentId { get; set; }
        public int InchargeId { get; set; }
        public int RoleId { get; set; }
        public int? ExplanationId { get; set; }
        public int? CompanyId { get; set; }
        public int? GradeId { get; set; }
        public decimal UnitPrice { get; set; }
        public string GradePoints { get; set; }
        public string GradeLowPoints { get; set; }
        public string GradeHighPoints { get; set; }

        public string GradeName { get; set; }
    }
}