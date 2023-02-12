using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class CommonMasterDAL:DbContext
    {
        public int CreateCommonMaster(CommonMaster commonMaster)
        {
            int result = 0;
            string query = $@"insert into CommonMasters(GradeId,SalaryIncreaseRate,OverWorkFixedTime,BonusReserveRatio,BonusReserveConstant,WelfareCostRatio,CreatedBy,CreatedDate) values(@gradeId,@salaryIncreaseRate,@overWorkFixedTime,@bonusReserveRatio,@bonusReserveConstant,@welfareCostRatio,@createdBy,@createdDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@gradeId", commonMaster.GradeId);
                cmd.Parameters.AddWithValue("@salaryIncreaseRate", commonMaster.SalaryIncreaseRate);
                cmd.Parameters.AddWithValue("@overWorkFixedTime", commonMaster.OverWorkFixedTime);
                cmd.Parameters.AddWithValue("@bonusReserveRatio", commonMaster.BonusReserveRatio);
                cmd.Parameters.AddWithValue("@bonusReserveConstant", commonMaster.BonusReserveConstant);
                cmd.Parameters.AddWithValue("@welfareCostRatio", commonMaster.WelfareCostRatio);
                cmd.Parameters.AddWithValue("@createdBy", commonMaster.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", commonMaster.CreatedDate);
                //cmd.Parameters.AddWithValue("@isActive", section.IsActive);
                try
                {
                    result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                }

                return result;
            }

        }
    }
}