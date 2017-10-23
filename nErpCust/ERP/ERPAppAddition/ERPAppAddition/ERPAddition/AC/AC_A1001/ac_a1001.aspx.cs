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
using ERPAppAddition.ERPAddition.AC.AC_A1001;

namespace ERPAppAddition.ERPAddition.AC.AC_A1001
{
    public partial class ac_a1001 : System.Web.UI.Page    
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_enc"].ConnectionString);
        string userid;
        
        SqlCommand sql_cmd = new SqlCommand();
        SqlCommand sql_cmd2 = new SqlCommand();
        SqlCommand sql_cmd3 = new SqlCommand();
        DataSet ds = new DataSet();
        int value;
        string setSQL = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {                   
                string yyyymm = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00");
                txtFrom.Text = yyyymm;
                txtTo.Text = yyyymm;
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

         

        protected void bt_retrieve_Click(object sender, EventArgs e)
        {

            string from = txtFrom.Text;
            String to = txtTo.Text;
            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            sql_cmd = conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = "USP_PROJ_COST "+ from +", "+ to;
            sql_cmd.CommandTimeout = 3000;

            sql_cmd2 = conn.CreateCommand();
            sql_cmd2.CommandType = CommandType.Text;
            sql_cmd2.CommandText = "USP_PROJ_COST2 " + from + ", " + to;
            sql_cmd2.CommandTimeout = 3000;

            sql_cmd3 = conn.CreateCommand();
            sql_cmd3.CommandType = CommandType.Text;
            sql_cmd3.CommandText = "USP_PROJ_COST3 " + from + ", " + to;
            sql_cmd3.CommandTimeout = 3000;
            
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                string report_nm = "rp_ac_a1001.rdlc";

                ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                ReportViewer1.LocalReport.DisplayName = "프로젝트별공사손익";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                /*두번째 grid*/
                SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);

                ReportDataSource rds2 = new ReportDataSource();
                rds2.Name = "DataSet2";
                rds2.Value = dt2;
                ReportViewer1.LocalReport.DataSources.Add(rds2);

                /*세번째 grid*/
                SqlDataAdapter da3 = new SqlDataAdapter(sql_cmd3);
                DataTable dt3 = new DataTable();
                da3.Fill(dt3);

                ReportDataSource rds3 = new ReportDataSource();
                rds3.Name = "DataSet3";
                rds3.Value = dt3;
                ReportViewer1.LocalReport.DataSources.Add(rds3);

                ReportViewer1.LocalReport.Refresh();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                Console.WriteLine("{0} Second exception caught.", ex);
            }
        }
    }
}