using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.DAL;

namespace CostAllocationApp.Controllers.Api
{
    public class SectionsController : ApiController
    {
        [HttpPost]
        public IHttpActionResult CreateSection(Section section)
        {
            //Section section = new Section();
            //section.SectionName = sectionName.ToString();
            //section.CreateBy = "";
            //.CreateDate = DateTime.Now;

            SectionDAL sectionDal = new SectionDAL();
            int result = sectionDal.CreateSection(section);
            if (result>0)
            {
                return Ok();
            }
            else
            {
                return BadRequest();
            }

            
        }
    }
}
