using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class CommonMastersController : ApiController
    {
        CommonMasterBLL _commonMasterBll = null;
        public CommonMastersController()
        {
            _commonMasterBll = new CommonMasterBLL();

        }
        [HttpPost]
        public IHttpActionResult CreateCommonMaster(CommonMaster commonMaster)
        {

            commonMaster.CreatedBy = "";
            commonMaster.CreatedDate = DateTime.Now;
            //section.IsActive = true;
            int result = _commonMasterBll.CreateCommonMaster(commonMaster);
            if (result > 0)
            {
                return Ok("Data Saved Successfully!");
            }
            else
            {
                return BadRequest("Something Went Wrong!!!");
            }
        }
        [HttpGet]
        public IHttpActionResult CommonMasters()
        {
            List<CommonMaster> commonMasters = _commonMasterBll.GetCommonMasters();
            return Ok(commonMasters);
        }
    }
}
