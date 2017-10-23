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
using Oracle.DataAccess.Client;



namespace ERPAppAddition.ERPAddition.AM.AM_AA1002
{
    public partial class AM_AA1002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);        
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;        

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ReportViewer1.Reset();
                /*달력셋*/
                setMonth();
                rbl_view_type.SelectedIndex = 0;

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

        private void setMonth()
        {
            cmb_yyyy.Text = DateTime.Now.Year.ToString();
            txt_mm.Text = DateTime.Now.Month.ToString("00");
        }

       protected void Load_btn_Click1(object sender, EventArgs e)
        {
            string tabIndexSql = ""; //tab에 따른 프로시저
            string tabIndexRdlc = ""; //tab에 따른 리포트
           switch(rbl_view_type.SelectedIndex)
           {
               case 0:
                   tabIndexSql = "dbo.USP_A_DAILY_AMT_DAILY";
                   tabIndexRdlc = "rp_am_aa1002_daily.rdlc";
                   break;
               case 1:
                   tabIndexSql = "dbo.USP_A_DAILY_AMT_MON";
                   tabIndexRdlc = "rp_am_aa1002_mon.rdlc";
                   break;
               case 2:
                   tabIndexSql = "dbo.USP_A_DAILY_AMT_BIZ";
                   tabIndexRdlc = "rp_am_aa1002_biz.rdlc";
                   break;
               default:
                   tabIndexSql = "dbo.USP_A_DAILY_AMT_DAILY";
                   tabIndexRdlc = "rp_am_aa1002_mon.rdlc";
                   break;
           }

            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = tabIndexSql;
            cmd.CommandTimeout = 3000;
           
           //파라미터 set
           SqlParameter param1 = new SqlParameter("@YYYY", SqlDbType.VarChar, 4);
           param1.Value = cmb_yyyy.Text;
           cmd.Parameters.Add(param1);
           //월일때는 파라미터 2(월)가 필요없음
           if (rbl_view_type.SelectedIndex != 1)
           {
               SqlParameter param2 = new SqlParameter("@MM", SqlDbType.VarChar, 2);
               param2.Value = txt_mm.Text;
               cmd.Parameters.Add(param2);
           }

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath(tabIndexRdlc);
                ReportViewer1.LocalReport.DisplayName = "일일운용자금" + txt_mm.Text + DateTime.Now.ToShortDateString();

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

       protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
       {
           if (rbl_view_type.SelectedIndex == 1)
           {
               txt_mm.Enabled = false;
               
           }
           else
           {
               txt_mm.Enabled = true;               
           }
       }
    }

}