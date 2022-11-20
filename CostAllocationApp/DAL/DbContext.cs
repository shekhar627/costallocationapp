using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Data.SqlClient;

namespace CostAllocationApp.DAL
{
    public class DbContext
    {
        private string _connectionString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;

        public SqlConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }
    }
}