using System;
using System.Collections.Generic;
using System.Web.Http;
using CostAllocationApp.Models;
using CostAllocationApp.BLL;

namespace CostAllocationApp.Controllers.Api
{
    public class ExplanationsController : ApiController
    {
        ExplanationsBLL explanationsBLL = null;
        public ExplanationsController()
        {
            explanationsBLL = new ExplanationsBLL();
        }

        [HttpPost]
        public IHttpActionResult CreateExplanation(Explanation explanation)
        {

            if (String.IsNullOrEmpty(explanation.ExplanationName))
            {
                return BadRequest("Explanation Name Required");
            }
            else
            {
                explanation.CreatedBy = "";
                explanation.CreatedDate = DateTime.Now;
                explanation.IsActive = true;


                int result = explanationsBLL.CreateExplanation(explanation);
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
        public IHttpActionResult Explanations()
        {
            List<Explanation> explanations = explanationsBLL.GetAllExplanations();
            return Ok(explanations);
        }

        [HttpDelete]
        public IHttpActionResult RemoveExplanations([FromUri] string explanationIds)
        {
            int result = 0;


            if (!String.IsNullOrEmpty(explanationIds))
            {
                string[] ids = explanationIds.Split(',');

                foreach (var item in ids)
                {
                    result += explanationsBLL.RemoveExplanations(Convert.ToInt32(item));
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