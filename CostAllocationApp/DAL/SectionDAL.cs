using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using CostAllocationApp.Models;

namespace CostAllocationApp.DAL
{
    public class SectionDAL:DbContext
    {

        public int CreateSection(Section section)
        {
            int result = 0;
            string query = $@"insert into Sections(Name,CreatedBy,CreatedDate,IsActive) values(@sectionName,@createdBy,@createdDate,@isActive)";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query,sqlConnection);
                cmd.Parameters.AddWithValue("@sectionName",section.SectionName);
                cmd.Parameters.AddWithValue("@createdBy", section.CreatedBy);
                cmd.Parameters.AddWithValue("@createdDate", section.CreatedDate);
                cmd.Parameters.AddWithValue("@isActive", section.IsActive);
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
            string query = "select * from Sections where isactive=1";
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
                            section.CreatedDate = Convert.ToDateTime(rdr["CreatedDate"]);
                            section.CreatedBy = rdr["CreatedBy"].ToString();

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

        public int RemoveSection(int sectionId)
        {
            int result = 0;
            string query = $@"update sections set isactive=0 where id=@id";
            using (SqlConnection sqlConnection = this.GetConnection())
            {
                sqlConnection.Open();
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@id", sectionId);
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