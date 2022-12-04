using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class EmployeeAssignment : Common
    {
        public int Id { get; set; }
        public string EmployeeName { get; set; }
        public int SectionId { get; set; }
        public int DepartmentId { get; set; }
        public int InchargeId { get; set; }
        public int RoleId { get; set; }
        public int ExplanationId { get; set; }
        public int CompanyId { get; set; }
        public decimal UnitPrice { get; set; }
        public int GradeId { get; set; }
        public string IsActive { get; set; }
        public string Remarks { get; set; }
        public int SubCode { get; set; }
    }
}