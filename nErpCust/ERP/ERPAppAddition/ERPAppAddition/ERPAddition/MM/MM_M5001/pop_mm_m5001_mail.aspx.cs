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
    public partial class pop_mm_m5001_mail : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        DataSet ds = new DataSet();
       
        
        protected void Page_Load(object sender, EventArgs e)
        {
            string sql;

            sql = "select UD_MINOR_CD 이름,UD_REFERENCE 메일주소" +
                 " from B_USER_DEFINED_MINOR " +
                 " where UD_MAJOR_CD='M0001'" +
                 " order by UD_MINOR_CD";



            DataTable dt = Execute_ERP(sql);

            FpSpread1.DataSource = ds;
            FpSpread1.DataBind();

            SpreadSet();
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
            string mail_add;
            mail_add = FpSpread1.Sheets[0].Cells[FpSpread1.ActiveSheetView.ActiveRow, 1].Text;
            // 위에서 받은 값을 해당 팝업페이지의 컨트롤 id에 셋팅한다.
            StringBuilder script = new StringBuilder();
            if (mail_add == "" || mail_add == null)
            {
                string msg = "alert(\"메일주소를 선택해 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", msg, true);
            }
            else
            {
                
                if ((Session["pgid"].ToString() == "mm_m5001") && (Session["popupid"].ToString() == "1"))
                {
                    script.Append("opener.document.getElementById(\"txt_mail_fr\").value = '" + mail_add + "';");
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



 