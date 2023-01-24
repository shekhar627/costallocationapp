using CostAllocationApp.DAL;
using System.Collections.Generic;
using CostAllocationApp.Models;

namespace CostAllocationApp.BLL
{
    public class AllocationBLL
    {
        AllocationDAL _allocationDAL= null;
        public AllocationBLL()
        {
            _allocationDAL = new AllocationDAL();
        }

        public int CreateAllocation(Allocation _allocation)
        {
            return _allocationDAL.CreateAllocation(_allocation);
        }

        public List<Allocation> GetAllocationList()
        {
            return _allocationDAL.GetAllocationList();
        }
        public int RemoveAllocation(int allocationId)
        {
            return _allocationDAL.RemoveAllocation(allocationId);
        }

        public bool CheckAllocation(string allocationName)
        {
            return _allocationDAL.CheckAllocation(allocationName);
        }

        public int GetAllocationCountWithEmployeeAsignment(int allocationId)
        {
            return _allocationDAL.GetAllocationCountWithEmployeeAsignment(allocationId);
        }
        public Allocation GetAllocationByAllocationId(int allocationId)
        {
            return _allocationDAL.GetAllocationByAllocationId(allocationId);
        }
    }
}