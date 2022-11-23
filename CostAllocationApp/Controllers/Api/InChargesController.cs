using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class InChargesController : ApiController
    {
        InChargeBLL inChargeBLL = null;
        public InChargesController()
        {
            inChargeBLL = new InChargeBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateInCharge(InCharge inCharge)
        {

            if (String.IsNullOrEmpty(inCharge.InChargeName))
            {
                return BadRequest("InCharge Name Required");
            }
            else
            {
                inCharge.CreatedBy = "";
                inCharge.CreatedDate = DateTime.Now;
                inCharge.IsActive = true;


                int result = inChargeBLL.CreateInCharge(inCharge);
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
        [HttpGet]
        public IHttpActionResult InCharges()
        {
            List<InCharge> inCharges = inChargeBLL.GetAllInCharges();
            return Ok(inCharges);
        }

        [HttpDelete]
        public IHttpActionResult RemoveInCharge([FromUri] string inChargeIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(inChargeIds))
            {
                string[] ids = inChargeIds.Split(',');

                foreach (var item in ids)
                {
                    result += inChargeBLL.RemoveInCharge(Convert.ToInt32(item));
                }

                if (result == ids.Length)
                {
                    return Ok("Data Removed Successfully!");
                }
                else
                {
                    return BadRequest("Something Went Wrong!!!");
                }
            }
            else
            {
                return BadRequest("Select InCharge Id!");
            }

        }

    }
}