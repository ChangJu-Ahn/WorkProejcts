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
using System.Collections;
using ERPAppAddition.ERPAddition.SM.sm_s5001;

namespace ERPAppAddition.ERPAddition.SM.sm_s5001
{
    public partial class sm_s50011 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();

        protected void Page_Load(object sender, EventArgs e)
        {
            //this.tb_yyyy.Attributes["onkeyPress"] = "if(event.keyCode == 13) {" +
            //    Page.GetPostBackEventReference(this.bt_retrieve) + "; return false}";

            if (!IsPostBack)
            {
                ReportViewer1.Reset();
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

        protected void bt_retrieve_Click(object sender, EventArgs e) //조회 버튼 클릭
        {
            string Procedure = string.Empty; // 구분자 사용을 위한 변수 (드롭박스로 상세조회, 집계조회를 하기 위함)
            string Report = string.Empty;
            string DataSheet = string.Empty;

            if (DropDownList1.Text == "Value_Detail") // 드롭박스에서 상세조회 선택 시 Value.Detail 값
            {
                Procedure = "dbo.USP_A_OPEN_AR_VIEW";
                Report = "rp_sm_5001.rdlc";
                DataSheet = "DataSet1";
            }
            else // 드롭박스에서 집계조회 선택 시 Value.Sum 값
            {
                Procedure = "dbo.USP_A_OPEN_AR_VIEW_SUM";
                Report = "rp_sm_5001_sum.rdlc";
                DataSheet = "DataSet2";
            }


            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = Procedure;
            cmd.CommandTimeout = 0;


            SqlParameter param1 = new SqlParameter("@YYYY", SqlDbType.VarChar, 4);
            SqlParameter param2 = new SqlParameter("@YYYY1", SqlDbType.VarChar, 4);
            SqlParameter param3 = new SqlParameter("@PAY_BP_CD", SqlDbType.VarChar, 20);

            string YYYY, PAY_BP_CD;
            YYYY = tb_yyyy.Text;
            PAY_BP_CD = tb_bp_cd.Text;
          

            param1.Value = tb_yyyy.Text; //기준년도
            if (tb_yyyy.Text == "" || tb_yyyy.Text == null)
                YYYY = "%";

            if (tb_yyyy.Text == "" || tb_yyyy.Text == null)
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 년 을 확인하세요')", true);
                return;
            }

                param2.Value = (Convert.ToInt32(tb_yyyy.Text) - 1).ToString();


                param3.Value = tb_bp_cd.Text; //거래처
                if (PAY_BP_CD == null || PAY_BP_CD == "")
                    PAY_BP_CD = "%";



                cmd.Parameters.Add(param1);
                cmd.Parameters.Add(param2);
                cmd.Parameters.Add(param3);

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(Report);
                    ReportViewer1.LocalReport.DisplayName = "미수금거래처관리(NEPES)" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = DataSheet;
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

           //거래처 조회 팝업
        protected void bt_bp_cd_Click1(object sender, EventArgs e) //거래처 조회 버튼 클릭
        {
            Response.Write("<script>window.open('pop_sm_s5001.aspx?pgid=sm_s5001&popupid=1','','top=100,left=100,width=800,height=600')</script>");
        }

    }
}