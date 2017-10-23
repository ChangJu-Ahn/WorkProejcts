using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;

public class cls_dbexe_erp
{

    SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
    SqlCommand cmd = new SqlCommand();
    SqlDataReader dr;
    int value;
    string sql;

    public int QueryExecute(string sql, string wk_type)
    {
        conn.Open();
        cmd = conn.CreateCommand();
        cmd.CommandType = CommandType.Text;
        cmd.CommandText = sql;

        try
        {
            //삭제시 기존 권한아이디에 프로그램이 연결되었는지 확인하기 위함.
            if (wk_type == "check")
                value = Convert.ToInt32(cmd.ExecuteScalar());
            else
                value = cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        conn.Close();
        return value;
    }

}
