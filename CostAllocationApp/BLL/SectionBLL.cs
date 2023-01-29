using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class SectionBLL
    {
        SectionDAL sectionDAL = null;
        public SectionBLL()
        {
            sectionDAL = new SectionDAL();
        }
        public List<Section> GetAllSections()
        {
            return sectionDAL.GetAllSections();
        }
        public int CreateSection(Section section)
        {
            return sectionDAL.CreateSection(section);
        }
        public int RemoveSection(int sectionId)
        {
            return sectionDAL.RemoveSection(sectionId);
        }

        public bool CheckSection(string sectionName)
        {
            return sectionDAL.CheckSection(sectionName);
        }

        public int GetSectionCountWithEmployeeAsignment(int sectionId)
        {
            return sectionDAL.GetSectionCountWithEmployeeAsignment(sectionId);
        }
        public Section GetSectionBySectionId(int sectionId)
        {
            return sectionDAL.GetSectionBySectionId(sectionId);
        }
    }
}