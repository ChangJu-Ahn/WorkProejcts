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


namespace ERPAppAddition.ERPAddition.EDU.Edu_GW
{
    public partial class e_edu_gw : System.Web.UI.Page
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

        protected void Button1_Click(object sender, EventArgs e)
        {

            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "USP_EDU_VIEW";     //SP명
            cmd.CommandTimeout = 0;

            SqlParameter param1 = new SqlParameter("@FR_YYYYMMDD", SqlDbType.VarChar, 8);
            SqlParameter param2 = new SqlParameter("@TO_YYYYMMDD", SqlDbType.VarChar, 8);
            SqlParameter param3 = new SqlParameter("@EDU_TYPE", SqlDbType.VarChar, 10);
            SqlParameter param4 = new SqlParameter("@SNO", SqlDbType.VarChar, 10);

            param1.Value = FR_YYYYMMDD.Text;
            param2.Value = TO_YYYYMMDD.Text;
            param3.Value = EDU_TYPE.SelectedValue;
            param4.Value = SNO.Text;

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_e_edu_gw.rdlc");
                ReportViewer1.LocalReport.DisplayName = "교육실적조회" + DateTime.Now.ToShortDateString();

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


