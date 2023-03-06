using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class ExportBLL
    {
        private ExportDAL _exportDal = null;
        private EmployeeAssignmentDAL _employeeAssignmentDAL = null;
        private CompanyBLL _companyBLL = null;
        public ExportBLL()
        {
            _exportDal = new ExportDAL();
            _employeeAssignmentDAL = new EmployeeAssignmentDAL();
            _companyBLL = new CompanyBLL();
        }

        public List<ForecastAssignmentViewModel> AssignmentsByAllocation(int departnmentId, int explanationId)
        {
            var _assignments = _exportDal.AssignmentsByAllocation(departnmentId, explanationId);

            if (_assignments.Count > 0)
            {
                foreach (var item in _assignments)
                {
                    item.forecasts = _employeeAssignmentDAL.GetForecastsByAssignmentId(item.Id);
                }
            }

            return _assignments;
        }

        public List<GradeUnitPriceType> GetGradeUnitPriceTypes(int gradeId, int departmentId, int year,int unitPriceTypeId)
        {
            return _exportDal.GetGradeUnitPriceTypes(gradeId, departmentId,year, unitPriceTypeId);
        }
    }
}