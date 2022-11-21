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

            if (String.IsNullOrEmpty(section.SectionName))
            {
                return BadRequest("Section Name Required");
            }
            else
            {
                section.CreateBy = "";
                section.CreateDate = DateTime.Now;
                section.IsActive = true;


                SectionDAL sectionDal = new SectionDAL();
                int result = sectionDal.CreateSection(section);
                if (result > 0)
                {
                    return Ok("Data Saved Successfully!");
                }
                else
                {
                    return BadRequest();
                }
            }



            
        }

        [HttpGet]
        public IHttpActionResult Sections()
        {
            SectionDAL sectionDal = new SectionDAL();
            List<Section> sections = sectionDal.GetAllSections();
            return Ok(sections);
        }

        [HttpPut]
        public IHttpActionResult EditSection(int sectionId)
        {
            SectionDAL sectionDal = new SectionDAL();
            int result = sectionDal.RemoveSection(sectionId);
            if (result > 0)
            {
                return Ok("Data Removed Successfully!");
            }
            else
            {
                return BadRequest();
            }
        }
    }
}
