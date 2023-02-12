using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.Models
{
    public class CommonMaster:Common
    {
        public int Id { get; set; }
        public int GradeId { get; set; }
        public string GradeName { get; set; }
        public decimal? SalaryIncreaseRate { get; set; }
        public decimal? OverWorkFixedTime { get; set; }
        public decimal? BonusReserveRatio { get; set; }
        public decimal? BonusReserveConstant { get; set; }
        public decimal? WelfareCostRatio { get; set; }
    }
}