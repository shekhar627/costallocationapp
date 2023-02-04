using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.ViewModels;
using CostAllocationApp.Models;


namespace CostAllocationApp.Dtos
{
    public class SalaryAssignmentDto
    {
        public Salary Salary { get; set; }
        public List<ForecastAssignmentViewModel> ForecastAssignmentViewModels { get; set; }
    }
}