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
    public class SectionsController : ApiController
    {
        SectionBLL sectionBLL = null;
        public SectionsController()
        {
            sectionBLL = new SectionBLL();
        }
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


                int result = sectionBLL.CreateSection(section);
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
        public IHttpActionResult Sections()
        {
            List<Section> sections = sectionBLL.GetAllSections();
            return Ok(sections);
        }

        [HttpDelete]
        public IHttpActionResult RemoveSection([FromUri] string sectionIds)
        {
            int result = 0;
            

            if (!String.IsNullOrEmpty(sectionIds))
            {
                string[] ids = sectionIds.Split(',');

                foreach (var item in ids)
                {
                    result += sectionBLL.RemoveSection(Convert.ToInt32(item));
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
                return BadRequest("Select Section Id!");
            }

        }
    }
}
