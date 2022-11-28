using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Dtos
{
    public class EmployeeAssignmentDTO
    {
        public string SectionId { get; set; }
        public string DepartmentId { get; set; }
        public string InchargeId { get; set; }
        public string RoleId { get; set; }
        public string ExplanationId { get; set; }
        public string CompanyId { get; set; }
        public string UnitPrice { get; set; }
        public string GradeId { get; set; }
    }
}