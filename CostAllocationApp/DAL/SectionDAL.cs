using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models; // temporary

namespace CostAllocationApp.DAL
{
    public class SectionDAL:DbContext
    {

        public int CreateSection(Section section)
        {
            int result = 0;
            string query = $@"insert into Sections(Name,CreatedBy,CreatedDate) values({section.SectionName},{section.CreateBy},{section.CreateDate})";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query,sqlConnection);
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