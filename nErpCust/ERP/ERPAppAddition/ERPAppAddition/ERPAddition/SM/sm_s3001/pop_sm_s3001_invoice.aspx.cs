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

namespace ERPAppAddition.ERPAddition.SM.sm_s3001
{
    public partial class pop_sm_s3001_invoice : System.Web.UI.Page
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
            string invoiceno,bill_cust_nm, ship_cust_nm, fr_dt, to_dt;
            invoiceno = tb_invoiceno.Text;
            bill_cust_nm = tb_bill_cust_nm.Text;
            ship_cust_nm = tb_ship_cust_nm.Text;
            fr_dt = tb_fr_dt.Text;
            to_dt = tb_to_dt.Text;

            if (invoiceno == null || invoiceno == "")
                invoiceno = "%";
            if (bill_cust_nm == null || bill_cust_nm == "")
                bill_cust_nm = "%";
            if (ship_cust_nm == null || ship_cust_nm == "")
                ship_cust_nm = "%";
            if (fr_dt == null || fr_dt == "")
                fr_dt = "00000000";
            if (to_dt == null || to_dt == "")
                to_dt = "99999999";


            sql = "select invoice_no 인보이스번호, invoice_dt 발행날자, bill_fr_cust_nm 수취인, ship_to_cust_nm 실물수령인 " +
                  "  from sm_invoice_hdr_nepes where invoice_dt >= '" + fr_dt + "' and  invoice_dt <= '" + to_dt + "'" +
                  "   and invoice_no like '" + invoiceno + "%' and bill_fr_cust_nm like '" + bill_cust_nm + "%' and ship_to_cust_nm like '" + ship_cust_nm + "%' ";
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
            string invoice_no;
            invoice_no = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 0].Text;
            // 위에서 받은 값을 해당 팝업페이지의 컨트롤 id에 셋팅한다.
            StringBuilder script = new StringBuilder();
            if (invoice_no == "" || invoice_no == null)
            {
                string msg = "alert(\"인보이스를 선택해 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", msg, true);
            }
            else
            {
                
                if ((Session["pgid"].ToString() == "sm_s3001") && (Session["popupid"].ToString() == "1"))
                {
                    script.Append("opener.document.getElementById(\"tb_invoice_no\").value = '" + invoice_no + "';");
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
    }
}