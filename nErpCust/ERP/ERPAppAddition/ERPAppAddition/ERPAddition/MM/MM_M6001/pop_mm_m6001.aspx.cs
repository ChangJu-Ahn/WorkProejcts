using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Data;
using System.Text;

namespace ERPAppAddition.ERPAddition.MM.MM_M6001
{
    public partial class pop_mm_m6001 : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        DataSet ds = new DataSet();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Session["pgid"] = Request.QueryString["pgid"];
                Session["popupid"] = Request.QueryString["popupid"];
            }
        }

        protected void btn_retrieve_Click(object sender, EventArgs e)
        {
            string sql;
            string bp_cd, bp_nm;
            bp_cd = tb_bp_cd.Text;
            bp_nm = tb_bp_nm.Text;


            if (bp_cd == null || bp_cd == "")
                bp_cd = "%";
            if (bp_nm == null || bp_nm == "")
                bp_nm = "%";


            sql = " Select  BP_CD 거래처코드,BP_NM 거래처명 From B_Biz_Partner Where BP_TYPE in ('S' ,'CS') " +
               " And BP_CD like '" + bp_cd + "%' and BP_nm like '" + bp_nm + "%' order by BP_CD";


            DataTable dt = Execute_ERP(sql);

            FpSpread1.DataSource = ds;
            FpSpread1.DataBind();

            SpreadSet();

        }
        private void SpreadSet()
        {
            FarPoint.Web.Spread.Column columnobj;
            int columncnt = FpSpread1.Columns.Count;
            columnobj = FpSpread1.ActiveSheetView.Columns[0, columncnt - 1];
            columnobj.Locked = true;
        }

        private DataTable Execute_ERP(string sql)
        {
            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;
            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                da.Fill(ds, "ds");
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
            conn_erp.Close();
            return dt;
        }

        protected void btn_ok_Click(object sender, EventArgs e)
        {
            string bp_cd;
            bp_cd = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 0].Text;
            // 위에서 받은 값을 해당 팝업페이지의 컨트롤 id에 셋팅한다.
            StringBuilder script = new StringBuilder();
            if (bp_cd == "" || bp_cd == null)
            {
                string msg = "alert(\"거래처를 선택해 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", msg, true);
            }
            else
            {

                if ((Session["pgid"].ToString() == "mm_m6001") && (Session["popupid"].ToString() == "1"))
                {
                    script.Append("opener.document.getElementById(\"tb_bp_cd\").value = '" + bp_cd + "';");
                    script.Append("window.close() ;");
                    ScriptManager.RegisterStartupScript(this, GetType(), "Script", script.ToString(), true);
                }

               

            }
        }

        protected void btn_cancel_Click(object sender, EventArgs e)
        {
            StringBuilder script = new StringBuilder();
            script.Append("window.close() ; ");

            ScriptManager.RegisterStartupScript(this, GetType(), "Script", script.ToString(), true);
        }

        protected void FpSpread1_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {

        }
    }
}
