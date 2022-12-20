using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CostAllocationApp.BLL
{
    public class InChargeBLL
    {
        InChargeDAL inChargeDAL = null;
        public InChargeBLL()
        {
            inChargeDAL = new InChargeDAL();
        }
        public int CreateInCharge(InCharge inCharge)
        {
            return inChargeDAL.CreateInCharge(inCharge);
        }
        public List<InCharge> GetAllInCharges()
        {
            return inChargeDAL.GetAllInCharges();
        }
        public int RemoveInCharge(int inChargeId)
        {
            return inChargeDAL.RemoveInCharge(inChargeId);
        }
        public bool CheckInCharge(string incharegeName)
        {
            return inChargeDAL.CheckInCharge(incharegeName);
        }
    }
}