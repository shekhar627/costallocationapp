using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.ViewModels
{
    public class GradeSalaryTypeViewModel : Common
    {
        public int Id { get; set; }
        public int GradeId { get; set; }
        public string GradeName { get; set; }
        public decimal GradeLowPoints { get; set; }
        public decimal GradeHighPoints { get; set; }
        public int DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public int SalaryTypeId { get; set; }
        public string SalaryTypeName { get; set; }
        public int Year { get; set; }
        public string AssignedGradeId { get; set; }
    }
}