using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.Models;

namespace CostAllocationApp.Dtos
{
    public class GradeCommonMasterDto
    {
        public Grade Grade { get; set; }
        public CommonMaster CommonMaster { get; set; }
    }
}