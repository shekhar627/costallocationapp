using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CostAllocationApp.DAL;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class CommonMasterBLL
    {
        CommonMasterDAL _commonMasterDal = null;
        public CommonMasterBLL()
        {
            _commonMasterDal = new CommonMasterDAL();
        }

        public int CreateCommonMaster(CommonMaster commonMaster)
        {
            return _commonMasterDal.CreateCommonMaster(commonMaster);
        }
        public List<CommonMaster> GetCommonMasters()
        {
            return _commonMasterDal.GetCommonMasters();
        }
    }
}