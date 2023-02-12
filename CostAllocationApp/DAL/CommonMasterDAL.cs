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
        public List<CommonMaster> GetCommonMasters()
        {
            List<CommonMaster> commonMasters = new List<CommonMaster>();
            string query = "";
            query = query + "select gm.Id,gm.GradeId,g.GradeName,gm.SalaryIncreaseRate,gm.OverWorkFixedTime ";
            query = query + "    ,gm.BonusReserveRatio,gm.BonusReserveConstant,gm.WelfareCostRatio ";
            query = query + "from CommonMasters gm ";
            query = query + "    join Grades g on gm.GradeId = g.Id ";

            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            CommonMaster commonMaster = new CommonMaster();
                            commonMaster.Id = Convert.ToInt32(rdr["Id"]);
                            commonMaster.GradeId = Convert.ToInt32(rdr["GradeId"]);                            
                            commonMaster.GradeName = rdr["GradeName"].ToString();
                            commonMaster.SalaryIncreaseRate = Convert.ToDecimal(rdr["SalaryIncreaseRate"]);
                            commonMaster.OverWorkFixedTime = Convert.ToDecimal(rdr["OverWorkFixedTime"]);
                            commonMaster.BonusReserveRatio = Convert.ToDecimal(rdr["BonusReserveRatio"]);
                            commonMaster.BonusReserveConstant = Convert.ToDecimal(rdr["BonusReserveConstant"]);
                            commonMaster.WelfareCostRatio = Convert.ToDecimal(rdr["WelfareCostRatio"]);
                            commonMasters.Add(commonMaster);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return commonMasters;
            }
        }

    }
}