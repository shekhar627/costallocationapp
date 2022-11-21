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
            string query = $@"insert into Sections(Name,CreatedBy,CreatedDate) values(@sectionName,@createBy,@createDate)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query,sqlConnection);
                cmd.Parameters.AddWithValue("@sectionName",section.SectionName);
                cmd.Parameters.AddWithValue("@createBy", section.CreateBy);
                cmd.Parameters.AddWithValue("@createDate", section.CreateDate);
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
        public List<Section> GetAllSections()
        {
            List<Section> sections = new List<Section>();
            string query = "select * from Sections";
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
                            Section section = new Section();
                            section.Id = Convert.ToInt32(rdr["Id"]);
                            section.SectionName = rdr["Name"].ToString();
                            section.CreateDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            section.CreateBy = rdr["CreatedBy"].ToString();

                            sections.Add(section);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                return sections;
            }
        }
    }
}