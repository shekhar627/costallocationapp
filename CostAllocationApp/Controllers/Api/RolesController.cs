using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class RolesController : ApiController
    {
        RoleBLL roleBLL = null;
        public RolesController()
        {
            roleBLL = new RoleBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateRole(Role role)
        {

            if (String.IsNullOrEmpty(role.RoleName))
            {
                return BadRequest("Role Name Required");
            }
            else
            {
                role.CreatedBy = "";
                role.CreatedDate = DateTime.Now;
                role.IsActive = true;


                int result = roleBLL.CreateRole(role);
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
            List<Role> roles = roleBLL.GetAllRoles();
            return Ok(roles);
        }

        [HttpDelete]
        public IHttpActionResult RemoveRoles([FromUri] string roleIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(roleIds))
            {
                string[] ids = roleIds.Split(',');

                foreach (var item in ids)
                {
                    result += roleBLL.RemoveRoles(Convert.ToInt32(item));
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