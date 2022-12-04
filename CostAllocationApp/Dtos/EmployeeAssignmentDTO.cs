using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Dtos
{
    public class EmployeeAssignmentDTO
    {
        public int Id { get; set; }
        public string EmployeeName { get; set; }
        public string SectionId { get; set; }
        public string DepartmentId { get; set; }
        public string InchargeId { get; set; }
        public string RoleId { get; set; }
        public string ExplanationId { get; set; }
        public string CompanyId { get; set; }
        public string UnitPrice { get; set; }
        public string GradeId { get; set; }
        public string Remarks { get; set; }
        public int SubCode { get; set; }


        public string[] Sections { get; set; }
        public string[] Departments { get; set; }
        public string[] Incharges { get; set; }
        public string[] Roles { get; set; }
        public string[] Explanations { get; set; }
        public string[] Companies { get; set; }
    }
}