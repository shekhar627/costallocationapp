using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace CostAllocationApp.ViewModels
{
    public abstract class BaseViewModel
    {
        public List<Department> Departments { get; set; }
    }
}