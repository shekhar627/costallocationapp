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
    public class AllocationController : ApiController
    {
        AllocationBLL _allocationBLL = null;

        public AllocationController()
        {
            _allocationBLL = new AllocationBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateAllocation(Allocation _allocation)
        {

            if (String.IsNullOrEmpty(_allocation.AllocationName))
            {
                return BadRequest("Allocation Name Required");
            }
            else
            {
                _allocation.CreatedBy = "";
                _allocation.CreatedDate = DateTime.Now;
                _allocation.IsActive = true;

                if (_allocationBLL.CheckAllocation(_allocation.AllocationName))
                {
                    return BadRequest("Allocation Already Exists!!!");
                }
                else
                {
                    int result = _allocationBLL.CreateAllocation(_allocation);
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
        public IHttpActionResult Alocations()
        {
            List<Allocation> _allocations = _allocationBLL.GetAllocationList();
            return Ok(_allocations);
        }

        [HttpDelete]
        public IHttpActionResult RemoveSection([FromUri] string allocationIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(allocationIds))
            {
                string[] ids = allocationIds.Split(',');

                foreach (var item in ids)
                {
                    result += _allocationBLL.RemoveAllocation(Convert.ToInt32(item));
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
                return BadRequest("Select Alocation Id!");
            }

        }

    }
}