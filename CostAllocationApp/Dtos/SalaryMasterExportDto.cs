using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;

namespace CostAllocationApp.Dtos
{
    public class SalaryMasterExportDto
    {
        public UnitPriceType UnitPriceType { get; set; }
        public List<GradeUnitPriceType> GradeSalaryTypes { get; set; }

    }
}