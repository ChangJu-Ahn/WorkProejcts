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

namespace ERPAppAddition.ERPAddition.AA.AA_A1001
{
    public partial class AA_A1001 : System.Web.UI.Page
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

                fr_yyyymmdd.Text = DateTime.Now.Year.ToString("0000") + "01";                
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void btn_view_Click(object sender, EventArgs e) //조회버튼
        {
            string Procedure = string.Empty;
            string Report = "AA_A1001.rdlc";
            string DataSheet = string.Empty;

            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "dbo.USP_DEPRECIATION_COST";
            cmd.CommandTimeout = 0;

            SqlParameter param1 = new SqlParameter("@S_DT", SqlDbType.VarChar, 6);
            param1.Value = fr_yyyymmdd.Text;

            cmd.Parameters.Add(param1);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath(Report);
                ReportViewer1.LocalReport.DisplayName = "Nepesway실적조회" + DateTime.Now.ToShortDateString();
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
            conn.Close();

        }
    }
}


