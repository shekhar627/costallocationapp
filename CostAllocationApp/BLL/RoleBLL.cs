using CostAllocationApp.DAL;
using CostAllocationApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
namespace CostAllocationApp.BLL
{
    public class RoleBLL
    {
        RoleDAL roleDAL = null;
        public RoleBLL()
        {
            roleDAL = new RoleDAL();
        }
        public int CreateRole(Role role)
        {
            return roleDAL.CreateRole(role);
        }
        public List<Role> GetAllRoles()
        {
            return roleDAL.GetAllRoles();
        }
        public int RemoveRoles(int roleIds)
        {
            return roleDAL.RemoveRoles(roleIds);
        }
        public bool CheckRole(string roleName)
        {
            return roleDAL.CheckRole(roleName);
        }
    }
}