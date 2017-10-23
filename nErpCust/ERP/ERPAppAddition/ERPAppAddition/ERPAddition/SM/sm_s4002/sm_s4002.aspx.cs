using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
//using System.Data.OracleClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;
namespace ERPAppAddition.ERPAddition.SM.sm_s4002
{
    public partial class sm_s4002 : System.Web.UI.Page
    {
       SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        SqlDataAdapter erp_sqlAdapter;
        DataSet ds = new DataSet();
        cls_dbexe_erp dbexe = new cls_dbexe_erp();
        string userid, db_name;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "") //사용자 ID값이 없다면 개발자 ID로할지 판단하기
                {
                    if (Request.QueryString["db"] == null || Request.QueryString["db"] == "") //DB없이 바로 실행할때 개발자용으로 적용
                        userid = "dev"; //erp에서 실행하지 않았을시 대비용
                    else // DB명이 있는데 사용자 ID가 없다면 이상하니 다시 접속하라는 메세지 보여줌
                    {
                        MessageBox.ShowMessage("잘못된 접근입니다. ERP접속 후 실행해주세요", this.Page);
                        this.Response.Redirect("../../Fail_Page.aspx");
                    }
                }
                else
                    userid = Request.QueryString["userid"];

                //MessageBox.ShowMessage(userid, this.Page);

                Session["User"] = userid;
              //  rbtn_work_type_SelectedIndexChanged(null, null);
                WebSiteCount();

            }
       
            //FpSpread_amt.Attributes.Add("onDataChanged", "ProfileSpread");
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_erp.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void ReportCreator(DataSet _dataSet, string sql, ReportViewer ReportViewer1, string _ReportName, string _ReportDataSourceName)
        {

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;

            DataSet ds = _dataSet;
            try
            {
                cmd_erp.CommandText = sql;
                dr_erp = cmd_erp.ExecuteReader();
                ds.Tables[0].Load(dr_erp);
                dr_erp.Close();
                ReportViewer1.LocalReport.ReportPath = Server.MapPath(_ReportName);

                ReportViewer1.LocalReport.DisplayName = "FCST 관리_"  +ddl_version.SelectedValue.ToString() + "_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                ReportViewer1.LocalReport.DataSources.Add(rds);
                ReportViewer1.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }

        }

     

        protected void btn_select_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            if (ddl_version.SelectedValue.ToString() == "-선택안함-" || ddl_version.SelectedValue.ToString() == null) 
            {
                MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                return;
            }
            if (tb_bas_yyyymm.Text == "" || tb_bas_yyyymm.Text == null)
            {
                MessageBox.ShowMessage("기준년월을 선택해주세요.", this.Page);
                return;
            }
           
           


            ReportViewer1.Reset();

            string sql = "SELECT distinct cust_nm,item_nm,item_gp,size,process_type,route,pkg_type,plan_mm,qty FROM S_FCST_QTY_IMPORT";//수량쿼리실행
            sql += " where bas_yyyymm = '" + tb_bas_yyyymm.Text + "'";//bas_yyyymm 과 선택한 기준년월은 동일
            sql += " and version_no =  '" + ddl_version.SelectedValue.ToString() + "'";//선택한 버전에 있는것만
            sql += " order by cust_nm,item_nm";//고객사, 디바이스 순 정렬
            ReportViewer1.Reset();
            ds_sm_s4002 dt1 = new ds_sm_s4002();

            ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4002.rdlc", "DataSet1");
        }

    }
}


