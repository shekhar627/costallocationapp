using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class UnitPriceTypeBLL
    {
        private UnitPriceTypeDAL _unitPriceTypeDAL = null;
        public UnitPriceTypeBLL()
        {
            _unitPriceTypeDAL = new UnitPriceTypeDAL();
        }

        public int CreateUnitPriceType(SalaryType unitPriceType)
        {
            return _unitPriceTypeDAL.CreateUnitPriceType(unitPriceType);
        }

        public bool CheckUnitPriceType(string unitPriceTypeName)
        {
            return _unitPriceTypeDAL.CheckUnitPriceType(unitPriceTypeName);
        }
        public List<SalaryType> GetAllSalaryTypes()
        {
            return _unitPriceTypeDAL.GetAllSalaryTypes();
        }
        public List<UnitPriceType> GetAllUnitPriceTypes()
        {
            return _unitPriceTypeDAL.GetAllUnitPriceTypes();
        }
        public UnitPriceType GetUnitPriceTypeById(int unitPriceTypeId)
        {
            return _unitPriceTypeDAL.GetUnitPriceTypeById(unitPriceTypeId);
        }
    }
}