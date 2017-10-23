using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using Microsoft.Reporting.WebForms;

namespace ERPAppAddition.ERPAddition.CM.CM_C7001
{
    public partial class cm_c7001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
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
            cmd.CommandText = "dbo.USP_C_BATCH_JOB_PROGRESS_VIEW";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 06);
            SqlParameter param2 = new SqlParameter("@work_step", SqlDbType.VarChar, 02);

            param1.Value = tb_yyyymm.Text;
            param2.Value = tb_workstep.Text;

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);


            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_cm_c7001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "원가진행현황리스트" + tb_yyyymm.Text + DateTime.Now.ToShortDateString();
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