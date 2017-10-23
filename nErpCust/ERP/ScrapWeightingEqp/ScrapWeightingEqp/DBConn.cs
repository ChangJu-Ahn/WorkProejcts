using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Web;
using System.Resources;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Windows.Forms;

namespace ScrapWeightingEqp
{
    public class DBConn
    {
        private static string _saveLogPath = System.Configuration.ConfigurationSettings.AppSettings["LOG_PATH"];
        static ScrapWeightingEqp.LogControl logCtrl = new LogControl(_saveLogPath, "Nepes_DB");

        #region 클래스 멤버변수
        private string _sConnInfo = string.Empty;
        SqlConnection sqlDBConn = null;
        #endregion

        #region 생성자
        public DBConn()
        {

        }
        public DBConn(string ConnInfo)
        {
            this._sConnInfo = ConnInfo;
        }
        #endregion

        #region 멤버변수 대입
        public string ConnectionString
        {
            get { return _sConnInfo; }
            set { _sConnInfo = value; }
        }
        #endregion

        #region DB연결 초기화(Initialize)
        public bool InitDBConn()
        {
            try
            {
                sqlDBConn = new SqlConnection(this._sConnInfo);
                sqlDBConn.Open();

                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region DB연결 종료(Close)
        public bool CloseDBConn()
        {
            try
            {
                if (sqlDBConn != null && sqlDBConn.State != ConnectionState.Closed)
                    sqlDBConn.Close();

                sqlDBConn = null;
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region DB쿼리 실행(SELECT 실행)
        public DataTable ExecuteQuery(string strQuery)
        {
            if (sqlDBConn == null || sqlDBConn.State != ConnectionState.Open)
                InitDBConn();

            SqlCommand sqlCmd = new SqlCommand(strQuery, sqlDBConn);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);

            try
            {
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);

                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message); //호출한 부분으로 예외 throw
            }
            finally
            {
                CloseDBConn();
            }
        }
        #endregion

        #region DB쿼리 실행(트랜잭션 쿼리 실행 / insert, delete, update)
        public int ExecuteNonQuery(string strQuery)
        {
            int nResult = 0;

            if (sqlDBConn == null || sqlDBConn.State != ConnectionState.Open)
                InitDBConn();

            SqlCommand sqlCmd = new SqlCommand(strQuery, sqlDBConn);

            try
            {
                nResult = sqlCmd.ExecuteNonQuery();
                return nResult;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                CloseDBConn();
            }
        }
        #endregion

        #region DB쿼리 실행(트랜잭션 쿼리 실행 / insert, delete, update)
        public DataSet ExecutePROCEDURE(string strProcedure, List<SqlParameter> param)
        {

            DataSet ds = new DataSet();
            if (sqlDBConn == null || sqlDBConn.State != ConnectionState.Open)
                InitDBConn();

            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd = sqlDBConn.CreateCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = strProcedure;
            sqlCmd.CommandTimeout = 3000;



            for (int i = 0; i < param.Count; i++)
            {
                sqlCmd.Parameters.Add(param[i]);

                if(param[i].Direction == ParameterDirection.Output)
                {
                    DataTable dt = new DataTable();
                    dt.TableName = "PARAM" +":"+ i;
                    dt.Columns.Add("MSG");
                    ds.Tables.Add(dt);
                }
            }


            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd);
                DataTable dtResult = new DataTable();
                da.Fill(dtResult);
                dtResult.TableName = "Result";
                if(ds.Tables.Count > 0)
                {
                    for(int i=0; i<ds.Tables.Count; i++)
                    {

                        int sParam = Convert.ToInt32(ds.Tables[i].TableName.Split(':')[1]);

                        ds.Tables[i].Rows.Add(new object[] { param[sParam].Value });

                    }
                }

                ds.Tables.Add(dtResult);

                return ds;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                CloseDBConn();
            }
        }
        #endregion

    }
}
