using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.ViewModels;
using CostAllocationApp.BLL;

namespace CostAllocationApp.BLL
{
    public class ExportBLL
    {
        private ExportDAL _exportDal = null;
        private EmployeeAssignmentDAL _employeeAssignmentDAL = null;
        public ExportBLL()
        {
            _exportDal = new ExportDAL();
            _employeeAssignmentDAL = new EmployeeAssignmentDAL();
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
    }
}