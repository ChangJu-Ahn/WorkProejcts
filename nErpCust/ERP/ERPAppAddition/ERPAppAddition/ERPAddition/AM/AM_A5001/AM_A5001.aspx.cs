using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.AM.AM_A5001;

namespace ERPAppAddition.ERPAddition.AM.AM_A5001
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string userid, db_name;

        protected void Page_Load(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            if (!Page.IsPostBack)
            {
                string sql = "";
                ds_am_a5001 dt1 = new ds_am_a5001();

                //if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
                //{
                    ////db_name = Request.QueryString["db"].ToString();
                    ////if (db_name.Length > 0)
                    ////{
                    ////    conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);

                    ////}
                    ////userid = Request.QueryString["userid"];
                    ////ReportCreator(dt1, sql, ReportViewer1, "rp_am_a5001.rdlc", "DataSet1");



                Session["FR_DT"]        = Request.QueryString["fr_dt"].ToString();     //상각년월 시작일
                Session["TO_DT"]        = Request.QueryString["to_dt"].ToString();     //상각년월 종료일
                Session["FROM_REG_DT"]  = Request.QueryString["reg_fr_dt"].ToString();
                Session["TO_REG_DT"]    = Request.QueryString["reg_to_dt"].ToString();
                Session["ASST_NO"]      = Request.QueryString["asst_no"].ToString();   //자산번호
                Session["DEPT_CD"]      = Request.QueryString["dept_cd"].ToString();   //관리부서
                Session["BizUnitCd"]    = Request.QueryString["bizunitcd"].ToString(); //사업부
                Session["DurYrsFg"]     = Request.QueryString["duryrsfg"].ToString(); //내용년수구분(C/T)
                Session["ACCT_CD"]      = Request.QueryString["acct_cd"].ToString();   //계정코드
                Session["FR_BizAreaCd"] = Request.QueryString["fr_bizareacd"].ToString(); //시작 사업장
                Session["TO_BizAreaCd"] = Request.QueryString["to_bizareacd"].ToString(); //종료 사업장

                     Load_btn_Click(null,null);
                //}
                 WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void Load_btn_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "A_ASSET_DEPR_MASTER_VIEW";
            cmd.CommandTimeout = 3000;

            SqlParameter param1 = new SqlParameter("@FROM_YYYYMM", SqlDbType.VarChar, 6);
            SqlParameter param2 = new SqlParameter("@TO_YYYYMM", SqlDbType.VarChar, 6);
            SqlParameter param3 = new SqlParameter("@FROM_REG_DT", SqlDbType.DateTime);
            SqlParameter param4 = new SqlParameter("@TO_REG_DT", SqlDbType.DateTime);
            SqlParameter param5 = new SqlParameter("@ASST_NO", SqlDbType.VarChar, 18);
            SqlParameter param6 = new SqlParameter("@DEPT_CD", SqlDbType.VarChar,10);
            SqlParameter param7 = new SqlParameter("@BP_CD", SqlDbType.VarChar, 10);
            SqlParameter param8 = new SqlParameter("@DUR_YRS_FG", SqlDbType.VarChar, 2);
            SqlParameter param9 = new SqlParameter("@ACCT_CD", SqlDbType.VarChar, 20);
            SqlParameter param10 = new SqlParameter("@FROM_BIZ_AREA_CD", SqlDbType.VarChar, 10);
            SqlParameter param11 = new SqlParameter("@TO_BIZ_AREA_CD", SqlDbType.VarChar, 10);
           


            param1.Value = Session["FR_DT"].ToString();       
            param2.Value = Session["TO_DT"].ToString();        
            param3.Value = Session["FROM_REG_DT"].ToString(); 
            param4.Value = Session["TO_REG_DT"].ToString(); 
            param5.Value = Session["ASST_NO"].ToString(); 
            param6.Value = Session["DEPT_CD"].ToString(); 
            param7.Value = Session["BizUnitCd"].ToString(); 
            param8.Value = Session["DurYrsFg"].ToString(); 
            param9.Value = Session["ACCT_CD"].ToString(); 
            param10.Value =Session["FR_BizAreaCd"].ToString();
            param11.Value = Session["TO_BizAreaCd"].ToString(); 

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);
            cmd.Parameters.Add(param5);
            cmd.Parameters.Add(param6);
            cmd.Parameters.Add(param7);
            cmd.Parameters.Add(param8); 
            cmd.Parameters.Add(param9); 
            cmd.Parameters.Add(param10);
            cmd.Parameters.Add(param11);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_a5001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "감가상각자산별조회" + DateTime.Now.ToShortDateString();

                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

                //UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        private void ReportCreator(DataSet _dataSet, string sql, ReportViewer ReportViewer1, string _ReportName, string _ReportDataSourceName)
        {

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = sql;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                ReportViewer1.LocalReport.ReportPath = Server.MapPath(_ReportName);

                ReportViewer1.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                ReportViewer1.LocalReport.DataSources.Add(rds);
                ReportViewer1.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

        }
    }
}


