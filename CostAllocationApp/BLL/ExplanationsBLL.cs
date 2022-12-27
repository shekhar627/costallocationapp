using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.BLL
{
    public class ExplanationsBLL
    {
        ExplanationDAL explanationDAL = null;
        public ExplanationsBLL()
        {
            explanationDAL = new ExplanationDAL();
        }
        public int CreateExplanation(Explanation explanation)
        {
            return explanationDAL.CreateExplanation(explanation);
        }
        public List<Explanation> GetAllExplanations()
        {
            return explanationDAL.GetAllExplanations();
        }
        public int RemoveExplanations(int explanationIds)
        {
            return explanationDAL.RemoveExplanations(explanationIds);
        }
        public Explanation GetSingleExplanationByExplanationId(int id)
        {
            return explanationDAL.GetSingleExplanationByExplanationId(id);
        }
        public int GetExplanationCountWithEmployeeAsignment(int explanationId)
        {
            return explanationDAL.GetExplanationCountWithEmployeeAsignment(explanationId);
        }
        public Explanation GetExplanationByExplanationId(int explanationId)
        {
            return explanationDAL.GetExplanationByExplanationId(explanationId);
        }
    }
}