using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class UnitPriceTypesController : ApiController
    {
        private UnitPriceTypeBLL _unitPriceTypeBLL = null;
        public UnitPriceTypesController()
        {
            _unitPriceTypeBLL = new UnitPriceTypeBLL();
        }
        // GET: UnitPriceTypes
        [HttpPost]
        public IHttpActionResult CreateUnitPriceType(SalaryType unitPriceType)
        {
            if (String.IsNullOrEmpty(unitPriceType.SalaryTypeName))
            {
                return BadRequest("Unit Price Type Name Required");
            }
            else
            {
                unitPriceType.CreatedBy = "";
                unitPriceType.CreatedDate = DateTime.Now;
                unitPriceType.IsActive = true;

                if (_unitPriceTypeBLL.CheckUnitPriceType(unitPriceType.SalaryTypeName))
                {
                    return BadRequest("Unit Price Type Name Already Exists!!!");
                }
                else
                {
                    int result = _unitPriceTypeBLL.CreateUnitPriceType(unitPriceType);
                    if (result > 0)
                    {
                        return Ok("Data Saved Successfully!");
                    }
                    else
                    {
                        return BadRequest("Something Went Wrong!!!");
                    }

                }
            }
        }

        [HttpGet]
        public IHttpActionResult UnitPriceTypes()
        {
            List<UnitPriceType> unitPriceTypes = _unitPriceTypeBLL.GetAllUnitPriceTypes();
            return Ok(unitPriceTypes);
        }
    }
}