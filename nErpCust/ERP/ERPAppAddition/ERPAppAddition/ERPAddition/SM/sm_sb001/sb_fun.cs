using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using ERPAppAddition.ERPAddition.SM.sm_sb001;

namespace ERPAppAddition.ERPAddition.SM.sm_sb001
{
    public class sb_fun
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];        

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_mail"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();

        public DataTable getData(string sql)
        {            
            SqlDataReader sql_dr;
            DataSet ds = new DataSet();

            DataTable retDt = new DataTable();            

            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = sql;
                
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds);
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();                

                retDt = ds.Tables[0];

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
            return retDt;
        }        

    }
}