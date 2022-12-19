using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class EmployeeAssignmentForecast
    {
        public int Id { get; set; }
        public string EmployeeName { get; set; }
        public string SectionId { get; set; }
        public string DepartmentId { get; set; }
        public string InchargeId { get; set; }
        public string RoleId { get; set; }
        public string ExplanationId { get; set; }
        public string CompanyId { get; set; }
        public decimal UnitPrice { get; set; }
        public int GradeId { get; set; }
        public string IsActive { get; set; }
        public string Remarks { get; set; }
        public int SubCode { get; set; }
    }
}