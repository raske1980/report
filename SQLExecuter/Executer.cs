using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;

namespace SQLExecuter
{
    public class Executer : IDisposable
    {
        private string constring = "Data Source=10.13.125.167;Initial Catalog=rIrCampus1.8Test;Persist Security Info=True;User ID=proof;Password=3371934proof;";
        SqlConnection sqlcon;

        public Executer()
        {
            OpenConnection();
        }

        public void OpenConnection()
        {
            sqlcon = new SqlConnection(constring);
            sqlcon.Open();
        }

        public SqlDataReader Execute(string command)
        {
            using (SqlCommand sqlCom = new SqlCommand(command, sqlcon))
            {
                sqlCom.CommandType = System.Data.CommandType.Text;
                return sqlCom.ExecuteReader();
            }
        }

        public void CloseConnection(SqlConnection sqlcon)
        {
            sqlcon.Close();
            sqlcon.Dispose();
        }

        public void Dispose()
        {
            CloseConnection(sqlcon);
        }
    }
}
