using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace T_IF_RCV_PROD_ORD_KO441
{
    public class DBConn
    {
        //Provider=SQLOLEDB.1;Data Source=MySQLServer;Initial Catalog=NORTHWIND;Integrated Security=SSPI
        private string sConnectionString = "";
        OleDbConnection oraConn = null;

        public string ConnectionString
        {
            get { return sConnectionString; }
            set { sConnectionString = value; }
        }

        public DBConn()
        {

        }

        public DBConn(string connectionString)
        {
            this.sConnectionString = connectionString;

            InitDBConn();
        }

        public bool InitDBConn()
        {
            try
            {
                oraConn = new OleDbConnection(this.sConnectionString);
                oraConn.Open();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public bool DisconnectDBConn()
        {
            if (oraConn != null && oraConn.State != ConnectionState.Closed)
            {
                oraConn.Close();
                oraConn = null;
            }

            return true;
        }

        public DataTable ExecuteQuery(string strQuery)
        {
            if (oraConn == null || oraConn.State != ConnectionState.Open)
                InitDBConn();

            OleDbCommand cmd = new OleDbCommand(strQuery, oraConn);

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(cmd);

            try
            {
                DataTable dataTable = new DataTable();

                dataAdapter.Fill(dataTable);

                dataAdapter.Dispose();
                oraConn.Close();

                return dataTable;
            }
            catch (Exception ex)
            {
                dataAdapter.Dispose();
                oraConn.Close();

                throw new Exception(ex.Message);
            }
        }

        public int ExecuteNonQuery(string strQuery)
        {
            int nResult = 0;

            if (oraConn == null)
            {
                InitDBConn();
            }
            else if (oraConn.State != ConnectionState.Open)
            {
                oraConn.Open();
            }

            OleDbCommand cmd = new OleDbCommand(strQuery, oraConn);
            //cmd.Transaction = oraConn.BeginTransaction();
            //cmd.Transaction.Begin();

            try
            {
                nResult = cmd.ExecuteNonQuery();

                //cmd.Transaction.Commit();

                oraConn.Close();
                return nResult;
            }
            catch (Exception ex)
            {
                //cmd.Transaction.Rollback();
                oraConn.Close();

                throw new Exception(ex.Message);
            }
        }
    }
}
