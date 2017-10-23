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

namespace ERPAppAddition.ERPAddition.SM.sm_sg001
{
    public partial class sm_sg001 : System.Web.UI.Page
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

        protected void bt_retrieve_Click(object sender, EventArgs e) //조회 버튼 클릭
        {
            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "dbo.USP_INVOICE_VIEW";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@FR_YYYYMMDD", SqlDbType.VarChar, 20);
            SqlParameter param2 = new SqlParameter("@TO_YYYYMMDD", SqlDbType.VarChar, 20);
            SqlParameter param3 = new SqlParameter("@INVOICE_NO",  SqlDbType.VarChar, 20);
            

            string sql;
            string FR_YYYYMMDD, TO_YYYYMMDD, INVOICE_NO;

            FR_YYYYMMDD = tb_fr_yyyymmdd.Text;
            TO_YYYYMMDD = tb_to_yyyymmdd.Text;
            INVOICE_NO = tb_invoice.Text;


            param1.Value = tb_fr_yyyymmdd.Text;
            if (FR_YYYYMMDD == null || FR_YYYYMMDD == "")
                FR_YYYYMMDD = "19900101";

            param2.Value = tb_to_yyyymmdd.Text;
            if (TO_YYYYMMDD == null || TO_YYYYMMDD == "")
                TO_YYYYMMDD = "29991231";

            param3.Value = tb_invoice.Text;
            if (INVOICE_NO == null || INVOICE_NO == "")
                INVOICE_NO = "%";

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
           


            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_sg001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "통관 인보이스 진행 현황 조회" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

               
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

     

    }
}

