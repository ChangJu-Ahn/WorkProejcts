using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace NEPES_CONTEXT
{
    /// <summary>
    /// Sql Parameter에 값을 담기위한 클래스 입니다.
    /// </summary>
    public class userSqlParams
    {
        public string KeyName { get; set; }
        public SqlDbType ColumnType { get; set; }
        public int ColunmSize { get; set; }
        public object Value { get; set; }
    }

    /// <summary>
    /// nepes DB Context
    /// </summary>
    public class NepesDBContext : IDisposable
    {
        public string connectionString { get; set; }

        /// <summary>
        /// 생성자입니다. 파라메터는 접속해야 할 Sql 접속정보 입니다
        /// </summary>
        public NepesDBContext(string value)
        {
            this.connectionString = value;
        }

        /// <summary>
        /// 데이터셋을 반환합니다.
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="userParams">쿼리에 사용되어야 할 Sql Parameter 구조체입니다.</param>
        public DataSet GetDataSet(string query, userSqlParams[] userParams)
        {
            System.Data.SqlClient.SqlParameter[] sqlParams = new System.Data.SqlClient.SqlParameter[userParams.Length];

            for (int i = 0; i < userParams.Length; i++)
            {
                sqlParams[i] = new SqlParameter("@" + userParams[i].KeyName, userParams[i].ColumnType, userParams[i].ColunmSize);
                sqlParams[i].Value = userParams[i].Value;
            }

            return SqlDataAccess.GetDataSet(query, sqlParams, connectionString);
        }

        /// <summary>
        /// 데이터 테이블을 반환합니다.
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="userParams">쿼리에 사용되어야 할 Sql Parameter 구조체입니다.</param>
        public DataTable GetDataTable(string query, userSqlParams[] userParams)
        {
            System.Data.SqlClient.SqlParameter[] sqlParams = new System.Data.SqlClient.SqlParameter[userParams.Length];

            for (int i = 0; i < userParams.Length; i++)
            {
                sqlParams[i] = new SqlParameter("@" + userParams[i].KeyName, userParams[i].ColumnType, userParams[i].ColunmSize);
                sqlParams[i].Value = userParams[i].Value;
            }

            return SqlDataAccess.GetDataTable(query, sqlParams, connectionString);
        }

        /// <summary>
        /// DML형식의 쿼리를 실행하여 실행완료 된 값을 숫자로 리턴합니다.(INSERT, UPDATE, DELETE)
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="userParams">쿼리에 사용되어야 할 Sql Parameter 구조체입니다.</param>
        public int ActionSqlQuery(string query, userSqlParams[] userParams)
        {
            System.Data.SqlClient.SqlParameter[] sqlParams = new System.Data.SqlClient.SqlParameter[userParams.Length];

            for (int i = 0; i < userParams.Length; i++)
            {
                sqlParams[i] = new SqlParameter("@" + userParams[i].KeyName, userParams[i].ColumnType, userParams[i].ColunmSize);
                sqlParams[i].Value = userParams[i].Value;
            }

            return SqlDataAccess.ExecuteNonQuery(query, sqlParams, connectionString);
        }


        public void Dispose() { }

    }

    class SqlDataAccess
    {
        /// <summary>
        /// 실제 DB에 접속하여 데이터 셋을 반환합니다.
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="userParams">쿼리에 사용되어야 할 Sql Parameter 구조체입니다.</param>
        /// <param name="connectionStr">접속되어야 할 DB정보입니다.</param>
        public static DataSet GetDataSet(string commandText, SqlParameter[] sqlParameters, string connectionStr)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataSet dsReturn = null;

            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = commandText;
                if (sqlParameters != null)
                {
                    foreach (SqlParameter param in sqlParameters)
                    {
                        AddParameter(cmd, param);
                    }
                }

                con = new SqlConnection(connectionStr);
                con.Open();

                cmd.Connection = con;

                dsReturn = new DataSet();
                da = new SqlDataAdapter(cmd);
                da.Fill(dsReturn);

                dsReturn.RemotingFormat = SerializationFormat.Binary;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                }
                if (cmd != null) cmd.Dispose();
                if (da != null) da.Dispose();
            }

            return dsReturn;
        }

        /// <summary>
        /// 실제 DB에 접속하여 데이터 테이블을 반환합니다.
        /// </summary>
        /// <param name="commandText">실행되어야 할 문장(쿼리)</param>
        /// <param name="sqlParameters">쿼리에 사용되어야 할 Sql Parameter 구조체입니다.</param>
        /// <param name="connectionStr">접속되어야 할 DB정보입니다.</param>
        public static DataTable GetDataTable(string commandText, SqlParameter[] sqlParameters, string connectionStr)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable dtReturn = null;

            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = commandText;
                if (sqlParameters != null)
                {
                    foreach (SqlParameter param in sqlParameters)
                    {
                        AddParameter(cmd, param);
                    }
                }

                con = new SqlConnection(connectionStr);
                con.Open();

                cmd.Connection = con;

                dtReturn = new DataTable();
                da = new SqlDataAdapter(cmd);
                da.Fill(dtReturn);

                dtReturn.RemotingFormat = SerializationFormat.Binary;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                }
                if (cmd != null) cmd.Dispose();
                if (da != null) da.Dispose();
            }

            return dtReturn;
        }

        /// <summary>
        /// SQL파라메터로 넘어온 값을 Sql Commend에 대입합니다.
        /// </summary>
        /// <param name="cmd">사용되어야 할 Commend입니다(Value Type : Reference Type)</param>
        /// <param name="param">Sql Parameter 입니다.</param>
        private static void AddParameter(SqlCommand cmd, SqlParameter param)
        {
            if ((param.Value == null) || ((param.Value.GetType().ToString().Equals("System.String")) && ((string)param.Value).Length == 0))
                param.Value = ""; //param.Value = System.DBNull.Value;
            cmd.Parameters.Add(param);
        }

        /// <summary>
        /// DMML 쿼리를 실행하여 결과 값을 반환합니다.
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="sqlParameters">쿼리 파라메터입니다(변수)</param>
        /// <param name="connectionStr">접속정보 입니다.</param>
        public static int ExecuteNonQuery(string query, SqlParameter[] sqlParameters, string connectionStr)
        {
            int nResult = 0;
            SqlConnection con = null;
            SqlCommand cmd = null;

            cmd = new SqlCommand();
            cmd.CommandText = query;
            if (sqlParameters != null)
                foreach (SqlParameter param in sqlParameters)
                    AddParameter(cmd, param);

            con = new SqlConnection(connectionStr);
            con.Open();

            cmd.Connection = con;

            try
            {
                nResult = cmd.ExecuteNonQuery();
                con.Close();
                return nResult;
            }
            catch (Exception ex)
            {
                con.Close();
                throw new Exception(ex.Message);
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }



    }

}
