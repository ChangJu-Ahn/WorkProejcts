﻿using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
//using System.Data.OleDb;
//using System.Data.OracleClient;
using System.IO;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;

namespace ERPAppAddition.ERPAddition.SM.sm_s9021
{
    public partial class sm_s9021 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        //string sql_cust_cd;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();        
        DataSet ds = new DataSet();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                DateTime dateTime = DateTime.Today.AddDays(-7);
                tb_fr_yyyymmdd.Text = dateTime.Year.ToString("0000") + dateTime.Month.ToString("00") + dateTime.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            if (tb_fr_yyyymmdd.Text == "" || tb_fr_yyyymmdd.Text == null || tb_to_yyyymmdd.Text == "" || tb_to_yyyymmdd.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"Date 를 선택해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }
            else
            {
                setGrid();
            }
        }

        private void setGrid()
        {
            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();                

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);

                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                    //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
                }
                sql_conn.Close();

                /*seq 가 a는 필수 항목 없는경우 조회 불가*/
                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s9021.rdlc");
                ReportViewer1.LocalReport.DisplayName = "반출데이터 조회" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = ds.Tables["DataSet1"];
                ReportViewer1.LocalReport.DataSources.Add(rds);
                
                ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
            }
        }

        private string getSQL()
        {
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	PRO_USP_DIST_01_S9021  '" + strFrom + "', '" + strTo + "' \n");
            return sbSQL.ToString();
        }
    }
}