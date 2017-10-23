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
using ERPAppAddition.ERPAddition.AM.AM_A4001;

namespace ERPAppAddition.ERPAddition.AM.AM_A4001
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        clsGlobalExec clsGE = new clsGlobalExec();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ReportViewer1.Reset();
                WebSiteCount();

                DataTable plant = clsGE.getData("SELECT CONVERT(VARCHAR(10), ALLC_DT, 120) AS ALLC_DT FROM A_CLS_ACCT_ITEM A, A_ALLC_HDR L WHERE A.CLS_NO = L.ALLC_NO AND A.ACCT_CD = '21100903' and ALLC_DT >= getdate()-121 GROUP BY ALLC_DT ORDER BY 1 DESC");               

                if (plant.Rows.Count > 0)
                {
                    DDL_DATE.DataTextField = "ALLC_DT";
                    DDL_DATE.DataValueField = "ALLC_DT";
                    DDL_DATE.DataSource = plant;
                    DDL_DATE.DataBind();
                }
                DDL_DATE.SelectedIndex = 0;

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


        protected void Load_btn_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "usp_bankmasterlist_view";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@gl_dt", SqlDbType.VarChar, 10);


            param1.Value = DDL_DATE.Text;

            cmd.Parameters.Add(param1);


            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_a4001_1.rdlc");
                ReportViewer1.LocalReport.DisplayName = "은행이체리스트(직원)" + DDL_DATE.Text + DateTime.Now.ToShortDateString();
               
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


       