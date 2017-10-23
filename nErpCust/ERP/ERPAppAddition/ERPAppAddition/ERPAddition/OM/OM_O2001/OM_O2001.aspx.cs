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
using ERPAppAddition.ERPAddition.OM.OM_O2001;

namespace ERPAppAddition.ERPAddition.OM.OM_O2001
{
    public partial class OM_O2001 : System.Web.UI.Page
    {
        SqlConnection conn= new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        protected void Page_Load(object sender, EventArgs e)
        {
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

        protected void btn_exe_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "dbo.usp_patentlist_view";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@apply_no", SqlDbType.VarChar, 20);
            SqlParameter param2 = new SqlParameter("@apply_fr_dt", SqlDbType.VarChar, 10);
            SqlParameter param3 = new SqlParameter("@apply_to_dt", SqlDbType.VarChar, 10);
            SqlParameter param4 = new SqlParameter("@apply_comp", SqlDbType.VarChar, 100);


            param1.Value = tb_apply_no.Text;
            param2.Value = tb_fr_dt.Text;
            param3.Value = tb_to_dt.Text;
            param4.Value = tb_apply_comp.Text;
            //if (tb_apply_no.Text == "" || tb_apply_no.Text == null)
            //    param1.Value = "%";

            //if (tb_fr_dt.Text == "" || tb_fr_dt.Text == null)
            //    param2.Value = "19990101";

            //if (tb_to_dt.Text == "" || tb_to_dt.Text == null)
            //    param3.Value = "29991231";

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);
            

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                
                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_om_o2001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "특허리스트" + tb_apply_no.Text + DateTime.Now.ToShortDateString();
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
    }
}