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

namespace OrderDraftMailSending
{
    public class DBConn
    {
        private static string _saveLogPath = System.Configuration.ConfigurationSettings.AppSettings["LOG_PATH"];
        static OrderDraftMailSending.LogControl logCtrl = new LogControl(_saveLogPath);

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

    }
}
