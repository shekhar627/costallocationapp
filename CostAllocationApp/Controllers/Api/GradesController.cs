using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class GradesController : ApiController
    {
        GradeBLL _gradeBLL = null;
        public GradesController()
        {
            _gradeBLL = new GradeBLL();
        }
        public IHttpActionResult CreateGrade(Grade grade)
        {

            if (String.IsNullOrEmpty(grade.GradeName))
            {
                return BadRequest("Grade Name Required");
            }
            else
            {
                grade.CreatedBy = "";
                grade.CreatedDate = DateTime.Now;
                //section.IsActive = true;

                if (_gradeBLL.CheckGrade(grade.GradeName))
                {
                    return BadRequest("Grade Already Exists!!!");
                }
                else
                {
                    int result = _gradeBLL.CreateGrade(grade);
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
        public IHttpActionResult Grades()
        {
            List<Grade> grades = _gradeBLL.GetAllGrade();
            return Ok(grades);
        }
    }
}