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

namespace ERPAppAddition.ERPAddition.POPUP
{
    public partial class pop_cust_cd : System.Web.UI.Page
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

        private void SpreadSet()
        {
            FarPoint.Web.Spread.Column columnobj;
            int columncnt = FpSpread1.Columns.Count;
            columnobj = FpSpread1.ActiveSheetView.Columns[0, columncnt - 1];
            columnobj.Locked = true;
        }

        protected void btn_retrieve_Click(object sender, EventArgs e)
        {
            string sql;
            string bp_cd, bp_nm;
            bp_cd = tb_cust_cd.Text;
            bp_nm = tb_cust_nm.Text;

            if (bp_cd == null || bp_cd == "")
                bp_cd = "%";
            if (bp_nm == null || bp_nm == "")
                bp_nm = "%";

            sql = "select BP_CD 거래처코드, BP_NM 거래처명, ADDR1 주소,ADDR1_ENG + ADDR2_ENG 영문주소, TEL_NO1 전화번호, FAX_NO FAX번호 from B_BIZ_PARTNER where bp_cd like '%" + bp_cd + "%' and bp_nm like '%" + bp_nm + "%' ";
            DataTable dt = Execute_ERP(sql);

            FpSpread1.DataSource = ds;
            FpSpread1.DataBind();

            SpreadSet();

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
            string bp_cd, bp_nm, addr_kr, addr_eng,tel_no, fax_no;
            bp_cd = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 0].Text;
            bp_nm = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 1].Text;
            addr_kr = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 2].Text;
            addr_eng = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 3].Text;
            tel_no = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 4].Text;
            fax_no = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 5].Text;
            // 위에서 받은 값을 해당 팝업페이지의 컨트롤 id에 셋팅한다.
            StringBuilder script = new StringBuilder();            
            ///script.Append ("<script type =\"text/javascript\">"); 
            //script.Append ("function AddData1() {");
            if (bp_cd == "" || bp_cd == null)
            {
                string msg = "alert(\"거래처를 선택해 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", msg, true);
            }
            else
            {
                //데이타 기본셋팅
                if (addr_kr == null || addr_kr == "") addr_kr = "주소없음";
                if (addr_eng == null || addr_eng == "") addr_eng = "주소없음";
                if (tel_no == null || tel_no == "") tel_no = "전화번호없음";
                if (fax_no == null || fax_no == "") fax_no = "FAX번호없음";

                if ((Session["pgid"].ToString() == "sm_s3001") && (Session["popupid"].ToString() == "1"))
                {
                    script.Append("opener.document.getElementById(\"hf_tb_ship_fr_cust_cd\").value = '" + bp_cd + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_fr_cust_nm\").value = '" + bp_nm + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_fr_add\").value = '" + addr_eng + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_fr_tel\").value = '" + tel_no + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_fr_fax\").value = '" + fax_no + "';");
                    script.Append("window.close();");
                    ScriptManager.RegisterStartupScript(this, GetType(), "Script", script.ToString(), true);
                }
                if ((Session["pgid"].ToString() == "sm_s3001") && (Session["popupid"].ToString() == "2"))
                {
                    script.Append("opener.document.getElementById(\"hf_tb_bill_to_cust_cd\").value = '" + bp_cd + "';");
                    script.Append("opener.document.getElementById(\"tb_bill_to_cust_nm\").value = '" + bp_nm + "';");
                    script.Append("opener.document.getElementById(\"tb_bill_to_add\").value = '" + addr_eng + "';");
                    script.Append("opener.document.getElementById(\"tb_bill_to_tel\").value = '" + tel_no + "';");
                    script.Append("opener.document.getElementById(\"tb_bill_to_fax\").value = '" + fax_no + "';");
                    script.Append("window.close();");
                    ScriptManager.RegisterStartupScript(this, GetType(), "Script", script.ToString(), true);
                }
                if ((Session["pgid"].ToString() == "sm_s3001") && (Session["popupid"].ToString() == "3"))
                {
                    script.Append("opener.document.getElementById(\"hf_tb_ship_to_cust_cd\").value = '" + bp_cd + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_to_cust_nm\").value = '" + bp_nm + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_to_add\").value = '" + addr_eng + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_to_tel\").value = '" + tel_no + "';");
                    script.Append("opener.document.getElementById(\"tb_ship_to_fax\").value = '" + fax_no + "';");
                    script.Append("window.close();");
                    ScriptManager.RegisterStartupScript(this, GetType(), "Script", script.ToString(), true);
                }
            }


        }
    }
}