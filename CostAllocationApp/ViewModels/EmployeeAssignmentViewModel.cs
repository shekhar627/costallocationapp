using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.ViewModels
{
    public class EmployeeAssignmentViewModel
    {
        public int SerialNumber { get; set; }
        public int Id { get; set; }
        public string EmployeeName { get; set; }
        public string SectionId { get; set; }
        public string SectionName { get; set; }
        public string DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public string InchargeId { get; set; }
        public string InchargeName { get; set; }
        public string RoleId { get; set; }
        public string RoleName { get; set; }
        public string ExplanationId { get; set; }
        public string ExplanationName { get; set; }
        public string CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string UnitPrice { get; set; }
        public string GradeId { get; set; }
        public string GradePoint { get; set; }
        public bool IsActive { get; set; }
        public string Remarks { get; set; }
        public int SubCode { get; set; }
        public bool MarkedAsRed { get; set; }
        public string EmployeeNameWithCodeRemarks { get; set; }
    }
}