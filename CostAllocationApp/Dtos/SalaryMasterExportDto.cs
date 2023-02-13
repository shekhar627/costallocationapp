using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;

namespace CostAllocationApp.Dtos
{
    public class SalaryMasterExportDto
    {
        public SalaryType SalaryType { get; set; }
        public List<GradeSalaryType> GradeSalaryTypes { get; set; }

    }
}