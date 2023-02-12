using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class GradeBLL
    {
        GradeDAL _gradeDAL = null;
        public GradeBLL()
        {
            _gradeDAL = new GradeDAL();
        }

        public int CreateGrade(Grade grade)
        {
            return _gradeDAL.CreateGrade(grade);
        }

        public bool CheckGrade(string gradeName)
        {
            return _gradeDAL.CheckGrade(gradeName);
        }
        public List<Grade> GetAllGrade()
        {
            return _gradeDAL.GetAllGrade();
        }
    }
}