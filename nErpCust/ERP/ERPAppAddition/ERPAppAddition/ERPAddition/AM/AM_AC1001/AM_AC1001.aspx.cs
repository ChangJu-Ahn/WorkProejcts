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
using ERPAppAddition.ERPAddition.AM.AM_AC1001;

namespace ERPAppAddition.ERPAddition.AM.AM_AC1001
{
    public partial class AM_AC1001 : System.Web.UI.Page
    {

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_display"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dReader_select;
        string id = "";
        string strSelQuery = ""; // 조회쿼리 변수

        //페이지 로딩 시 선언 및 호출 함수
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["id"] != null)
            {
                id = Request.QueryString["id"];
            }
            else
                id = "";

            switch (rbl_view_type.SelectedValue)
            {
                case "A": // 일 자금실적조회 쿼리
                    strSelQuery = " SELECT * FROM A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "' ORDER BY YYYY, MM ";
                    break;

                case "B": // 월 자금실적조회 쿼리
                    strSelQuery = " USP_A_DAILY_AMT_SEL_MONTH '" + txt_yyyy.Text + "' ";
                    break;

                case "C": // 계획대비 실적분석 쿼리
                    strSelQuery = " SELECT * FROM A_DAILY_AMT_PLAN WHERE YYYY = '" + txt_yyyy.Text + "' ORDER BY YYYY ";
                    break;
            }

            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        //메뉴 숨김 컨트롤 함수
        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (rbl_view_type.SelectedValue)
            {
                case "A": // 일 자금실적 조회
                    table.Visible = true;
                    ld_yyyy.Visible = true;
                    txt_yyyy.Visible = true;
                    ld_mm.Visible = true;
                    txt_mm.Visible = true;
                    td.Visible = true;
                    Select_Button.Visible = true;
                    break;

                case "B": // 월 자금실적 조회
                    table.Visible = true;
                    ld_yyyy.Visible = true;
                    txt_yyyy.Visible = true;
                    ld_mm.Visible = false;
                    txt_mm.Visible = false;
                    td.Visible = false;
                    Select_Button.Visible = true;
                    break;

                case "C": // 계획대비 실적분석 조회
                    table.Visible = true;
                    ld_yyyy.Visible = true;
                    txt_yyyy.Visible = true;
                    ld_mm.Visible = false;
                    txt_mm.Visible = false;
                    td.Visible = false;
                    Select_Button.Visible = true;
                    break;
            }
        }

        // 조회버튼 함수
        protected void btn_Select_Click(object sender, EventArgs e)
        {
            conn.Open();
            string Report = string.Empty;
            ReportViewer1.Reset();

            if (rbl_view_type.SelectedValue == "A")
            {
                if (txt_yyyy == null || txt_yyyy.Text.Equals("") || txt_mm == null || txt_mm.Text.Equals(""))
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('조회날짜를 확인하세요.')", true);
                    return;
                }
            }

            else
            {
                if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('조회날짜를 확인하세요.')", true);
                    return;
                }
            }

            switch (rbl_view_type.SelectedValue)
            {
                case "A":
                    Report = "DAY_REPORT.RDLC";
                    break;

                case "B":
                    Report = "MON_REPORT.RDLC";
                    break;

                case "C":
                    Report = "PLAN_REPORT.RDLC";
                    break;
            }

            cmd.Connection = conn;
            cmd.CommandText = strSelQuery;
            dReader_select = cmd.ExecuteReader();

            if (dReader_select.Read())
            {
                Display_AMT dt1 = new Display_AMT();
                ReportViewer1.Reset();
                conn.Close();

                ReportCreator(dt1, strSelQuery, ReportViewer1, Report, "DataSet1");
            }

            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜의 데이터가 없습니다.')", true);
            }

            conn.Close();

        }

        // 레포트뷰어 컨트롤 함수
        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;

            try
            {
                cmd.CommandText = _Query;
                dReader_select = cmd.ExecuteReader();

                switch (rbl_view_type.SelectedValue)
                {
                    case "A":
                        ds.Tables[0].Load(dReader_select);
                        break;

                    case "B":
                        ds.Tables[2].Load(dReader_select);
                        break;

                    case "C":
                        ds.Tables[1].Load(dReader_select);
                        break;
                }

                dReader_select.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                if (rbl_view_type.SelectedValue == "A")
                {
                    _reportViewer.LocalReport.DisplayName = "REPORT_Display_일자금조회_" + DateTime.Now.ToShortDateString();
                }
                else if (rbl_view_type.SelectedValue == "B")
                {
                    _reportViewer.LocalReport.DisplayName = "REPORT_Display_월자금조회_" + DateTime.Now.ToShortDateString();
                }
                else
                {
                    _reportViewer.LocalReport.DisplayName = "REPORT_Display_실적분석_" + DateTime.Now.ToShortDateString();
                }

                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;

                switch (rbl_view_type.SelectedValue)
                {
                    case "A":
                        rds.Value = ds.Tables[0];
                        break;

                    case "B":
                        rds.Value = ds.Tables[2];
                        break;

                    case "C":
                        rds.Value = ds.Tables[1];
                        break;
                }

                _reportViewer.LocalReport.DataSources.Add(rds);
                _reportViewer.LocalReport.Refresh();
            }

            catch
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('애러가 발생되었습니다. 익스플로어 종료 후 재실행 부탁 드립니다.')", true);
            }

            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }


    }
}