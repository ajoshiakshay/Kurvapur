using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;

namespace KurvapurSevaApp.DBConn
{
    class DBConnection
    {
        private SqlConnection conn;
        private string sDBConn = ConfigurationSettings.AppSettings["DBConn"];

        public SqlConnection getSQLConn() 
        {
            conn = new SqlConnection(sDBConn);
            return conn;
        }

        public void openSQLConn()
        {
            conn.Open();
        }

        public void closeSQLConn()
        {
            conn.Close();
        }
    }
}
