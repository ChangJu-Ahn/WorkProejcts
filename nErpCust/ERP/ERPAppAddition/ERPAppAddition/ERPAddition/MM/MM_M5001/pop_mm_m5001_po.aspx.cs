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

namespace ERPAppAddition.ERPAddition.MM.MM_M5001
{
    public partial class pop_mm_m5001_po : System.Web.UI.Page
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
        
        protected void btn_retrieve_Click(object sender, EventArgs e) //조회
        {
            string sql;
            string po_no,po_fr_dt, po_to_dt,attn_nm,po_person_nm;
            po_no = tb_po_no.Text;
            po_fr_dt = tb_po_fr_dt.Text;
            po_to_dt = tb_po_to_dt.Text;
            attn_nm = tb_attn_nm.Text;
            po_person_nm = tb_po_person_nm.Text;

            if (po_no == null || po_no == "")
                po_no = "%";
            if (po_fr_dt == null || po_fr_dt == "")
                po_fr_dt = "20010101";
            if (po_to_dt == null || po_to_dt == "")
                po_to_dt = DateTime.Now.ToString("yyyyMMdd");
            if (attn_nm == null || attn_nm == "")
                attn_nm = "%";
            if (po_person_nm == null || po_person_nm == "")
                po_person_nm = "%";

           sql = "select distinct a.PO_NO 발주번호, a.PO_DT 발주날짜 , a.BP_CD 공급처코드 ,d.BP_FULL_NM 공급처명,a.PUR_GRP 구매그룹코드,c.PUR_GRP_NM 구매그룹명"+
                 ",sum(b.PO_QTY)수량, CONVERT(VARCHAR(50), CAST( a.TOT_PO_LOC_AMT AS MONEY),1) 금액" +
                 " from m_pur_ord_hdr a (nolock) INNER JOIN m_pur_ord_dtl b (nolock) on a.po_no=b.po_no" +
                 " INNER JOIN b_pur_grp c on a.PUR_GRP=c.PUR_GRP INNER JOIN b_biz_partner d on a.BP_CD=d.BP_CD" +
                 " INNER JOIN B_ITEM F ON B.ITEM_CD=F.ITEM_CD    " +
                 " INNER JOIN T_IF_SND_PO_HDR_KO441 G on A.PO_NO= G.PO_NO and A.BP_CD=G.BP_CD"+
                 " where G.APPROVE_STATUS='7' and a.po_dt >= '" + po_fr_dt + "' and  a.po_dt <= '" + po_to_dt + "'" +
                 " and PLANT_CD not in ('P03','P04','P10')  and a.po_no like '" + po_no + "%' and BP_FULL_NM like '" + attn_nm + "%' and c.PUR_GRP_NM like '" + po_person_nm + "%' " +
                 " group by a.PO_NO,a.PO_DT,a.BP_CD,d.BP_FULL_NM,a.PUR_GRP,c.PUR_GRP_NM ,a.TOT_PO_LOC_AMT"+
                 "  order by a.PO_NO";  

     
             
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
            string po_no;
            po_no = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 0].Text;
            // 위에서 받은 값을 해당 팝업페이지의 컨트롤 id에 셋팅한다.
            StringBuilder script = new StringBuilder();
            if (po_no == "" || po_no == null)
            {
                string msg = "alert(\"발주번호를 선택해 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", msg, true);
            }
            else
            {
                
                if ((Session["pgid"].ToString() == "mm_m5001") && (Session["popupid"].ToString() == "1"))
                {
                    script.Append("opener.document.getElementById(\"tb_po_no\").value = '" + po_no + "';");
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

