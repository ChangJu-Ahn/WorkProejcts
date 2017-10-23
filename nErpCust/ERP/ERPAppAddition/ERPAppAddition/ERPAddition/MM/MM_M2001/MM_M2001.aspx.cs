using System;
using System.Data;
using System.Text;
//using System.Data.SqlClient;
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
using System.Drawing;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.QueryExe;

namespace ERPAppAddition.ERPAddition.MM.MM_M2001
{
    public partial class MM_M2001 : System.Web.UI.Page
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
                rbtn_work_type_SelectedIndexChanged(null, null);

                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_erp.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void btn_exe_Click(object sender, EventArgs e)
        {
            string bas_yyyymm = tb_bas_yyyymm.Text;
            string version_no = ddl_version.SelectedValue.ToString();
            string work_yyyymmdd = tb_work_yyyymmdd.Text;
            ReportViewer1.Reset();
            if (bas_yyyymm == null || bas_yyyymm == "")
            {
                MessageBox.ShowMessage("기준년월을 입력해주세요.", this.Page);
            }
            else if (work_yyyymmdd == null || work_yyyymmdd == "")
                MessageBox.ShowMessage("작업일자를 입력해주세요.", this.Page);
            else
            {
                conn_erp.Open();
                cmd_erp = conn_erp.CreateCommand();
                cmd_erp.CommandType = CommandType.StoredProcedure;
                cmd_erp.CommandText = "dbo.USP_CREATE_MRP_USAGE_EXE";
                cmd_erp.CommandTimeout = 3000;
                SqlParameter param1 = new SqlParameter("@view_type", SqlDbType.VarChar, 10);
                SqlParameter param2 = new SqlParameter("@version_no", SqlDbType.VarChar, 05);
                SqlParameter param3 = new SqlParameter("@bas_yyyymm", SqlDbType.VarChar, 06);
                SqlParameter param4 = new SqlParameter("@prnt_item_cd", SqlDbType.VarChar, 30);
                SqlParameter param5 = new SqlParameter("@CHILD_ITEM_CD", SqlDbType.VarChar, 30);

                param1.Value = "TOTAL"; //DETAIL도 있는데 일단 전체로
                param2.Value = version_no;
                param3.Value = bas_yyyymm;
                param4.Value = "%";
                param5.Value = "%";

                cmd_erp.Parameters.Add(param1);
                cmd_erp.Parameters.Add(param2);
                cmd_erp.Parameters.Add(param3);
                cmd_erp.Parameters.Add(param4);
                cmd_erp.Parameters.Add(param5);

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    string report_nm;
                    report_nm = "rv_mm_m2001.rdlc";

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                    ReportViewer1.LocalReport.DisplayName = work_yyyymmdd + "_자재소요량_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dt;
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    ReportViewer1.LocalReport.Refresh();
                    if (dt.Rows.Count > 0)
                        MessageBox.ShowMessage("계산되었습니다.", this.Page);
                    else
                        MessageBox.ShowMessage("계산된 데이타가 없습니다.", this.Page);

                }
                catch (Exception ex)
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                    MessageBox.ShowMessage("계산시 오류가 발생했습니다.", this.Page);
                }
            }
        }

        protected void rbtn_work_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbtn_work_type.SelectedValue == "create")
            {
                btn_view.Visible = false;
                btn_exe.Visible = true;
                btn_save.Visible = true;
                Btn_fcst_view.Visible = false; // fcst 조회 버튼
            }
            else
            {
                btn_view.Visible = true;
                btn_exe.Visible = false;
                btn_save.Visible = false;
                Btn_fcst_view.Visible = true; // fcst 조회 버튼
            }
            ReportViewer1.Reset();
        }

        protected void btn_save_Click(object sender, EventArgs e)
        {
            string bas_yyyymm = tb_bas_yyyymm.Text;
            string version_no = ddl_version.SelectedValue.ToString();
            string work_yyyymmdd = tb_work_yyyymmdd.Text;
            ReportViewer1.Reset();
            if (bas_yyyymm == null || bas_yyyymm == "")
            {
                MessageBox.ShowMessage("기준년월을 입력해주세요.", this.Page);
            }
            else if (work_yyyymmdd == null || work_yyyymmdd == "")
                MessageBox.ShowMessage("작업일자를 입력해주세요.", this.Page);
            else
            {
                conn_erp.Open();
                cmd_erp = conn_erp.CreateCommand();
                cmd_erp.CommandType = CommandType.StoredProcedure;
                cmd_erp.CommandText = "dbo.USP_CREATE_MRP_USAGE_INSERT";
                cmd_erp.CommandTimeout = 3000;
                SqlParameter param1 = new SqlParameter("@view_type", SqlDbType.VarChar, 10);
                SqlParameter param2 = new SqlParameter("@version_no", SqlDbType.VarChar, 05);
                SqlParameter param3 = new SqlParameter("@bas_yyyymm", SqlDbType.VarChar, 06);
                SqlParameter param4 = new SqlParameter("@prnt_item_cd", SqlDbType.VarChar, 30);
                SqlParameter param5 = new SqlParameter("@CHILD_ITEM_CD", SqlDbType.VarChar, 30);
                SqlParameter param6 = new SqlParameter("@work_dt", SqlDbType.VarChar, 08);
                SqlParameter param7 = new SqlParameter("@userid", SqlDbType.VarChar, 20);
                param1.Value = "TOTAL"; //DETAIL도 있는데 일단 전체로
                param2.Value = version_no;
                param3.Value = bas_yyyymm;
                param4.Value = "%";
                param5.Value = "%";
                param6.Value = work_yyyymmdd;
                param7.Value = Session["User"];
                cmd_erp.Parameters.Add(param1);
                cmd_erp.Parameters.Add(param2);
                cmd_erp.Parameters.Add(param3);
                cmd_erp.Parameters.Add(param4);
                cmd_erp.Parameters.Add(param5);
                cmd_erp.Parameters.Add(param6);
                cmd_erp.Parameters.Add(param7);

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    string report_nm;
                    report_nm = "rv_mm_m2001.rdlc";

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                    ReportViewer1.LocalReport.DisplayName = work_yyyymmdd + "_자재소요량_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dt;
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    ReportViewer1.LocalReport.Refresh();
                    if (dt.Rows.Count > 0)
                        MessageBox.ShowMessage("저장되었습니다.", this.Page);
                    else
                        MessageBox.ShowMessage("저장된 데이타가 없습니다.", this.Page);
                }
                catch (Exception ex)
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                    MessageBox.ShowMessage("저장시 오류가 발생했습니다.", this.Page);
                }
            }
        }

        protected void btn_view_Click(object sender, EventArgs e)
        {
            string bas_yyyymm = tb_bas_yyyymm.Text;
            string version_no = ddl_version.SelectedValue.ToString();
            string work_yyyymmdd = tb_work_yyyymmdd.Text;

            ReportViewer1.Reset();
            if (ddl_version.SelectedValue.ToString() == "-선택안함-" || ddl_version.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                return;
            }
            if (bas_yyyymm == null || bas_yyyymm == "")
            {
                MessageBox.ShowMessage("기준년월을 입력해주세요.", this.Page);
            }
            else if (work_yyyymmdd == null || work_yyyymmdd == "")
                MessageBox.ShowMessage("작업일자를 입력해주세요.", this.Page);
            else
            {
                conn_erp.Open();
                cmd_erp = conn_erp.CreateCommand();
                cmd_erp.CommandType = CommandType.StoredProcedure;
                cmd_erp.CommandText = "dbo.USP_CREATE_MRP_USAGE_VIEW";
                cmd_erp.CommandTimeout = 3000;
                SqlParameter param1 = new SqlParameter("@view_type", SqlDbType.VarChar, 10);
                SqlParameter param2 = new SqlParameter("@version_no", SqlDbType.VarChar, 05);
                SqlParameter param3 = new SqlParameter("@bas_yyyymm", SqlDbType.VarChar, 06);
                SqlParameter param4 = new SqlParameter("@prnt_item_cd", SqlDbType.VarChar, 30);
                SqlParameter param5 = new SqlParameter("@CHILD_ITEM_CD", SqlDbType.VarChar, 30);
                SqlParameter param6 = new SqlParameter("@work_dt", SqlDbType.VarChar, 08);
                param1.Value = "TOTAL"; //DETAIL도 있는데 일단 전체로
                param2.Value = version_no;
                param3.Value = bas_yyyymm;
                param4.Value = "%";
                param5.Value = "%";
                param6.Value = work_yyyymmdd;

                cmd_erp.Parameters.Add(param1);
                cmd_erp.Parameters.Add(param2);
                cmd_erp.Parameters.Add(param3);
                cmd_erp.Parameters.Add(param4);
                cmd_erp.Parameters.Add(param5);
                cmd_erp.Parameters.Add(param6);


                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    string report_nm;
                    report_nm = "rv_mm_m2001.rdlc";

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                    ReportViewer1.LocalReport.DisplayName = work_yyyymmdd + "_자재소요량_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dt;
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    ReportViewer1.LocalReport.Refresh();
                    if (dt.Rows.Count > 0)
                        MessageBox.ShowMessage("조회되었습니다.", this.Page);
                    else
                        MessageBox.ShowMessage("데이타가 없습니다.", this.Page);
                }
                catch (Exception ex)
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                    MessageBox.ShowMessage("조회시 오류가 발생했습니다.", this.Page);
                }
            }
        }

        protected void Btn_fcst_view_Click(object sender, EventArgs e) //fcst 조회
        {

            StringBuilder strBuilder = new StringBuilder();

            strBuilder.Append("<script language='javascript'>");
            strBuilder.Append("w=1600;h=650;");
            strBuilder.Append("x=Math.floor( (screen.availWidth-(w+12))/2 );y=Math.floor( (screen.availHeight-(h+30))/2 );");
            strBuilder.Append("window.open('../../sm/sm_s4002/sm_s4002.aspx', '',");
            strBuilder.Append("'height='+h+',width='+w+',top='+y+',left='+x+',scrollbars=yes,resizable=no');");
            strBuilder.Append("</script>");

            if (!ClientScript.IsClientScriptBlockRegistered("PopupScript"))
            {
                ClientScript.RegisterClientScriptBlock(this.GetType(), "PopupScript", strBuilder.ToString());
            }
             
            }

         
         }
}
