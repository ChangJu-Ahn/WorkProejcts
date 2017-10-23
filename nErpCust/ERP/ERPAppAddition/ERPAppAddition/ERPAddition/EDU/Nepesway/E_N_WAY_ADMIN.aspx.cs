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


namespace ERPAppAddition.ERPAddition.EDU.Nepesway
{
    public partial class E_N_WAY_ADMIN : System.Web.UI.Page
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

        protected void btn_view_Click(object sender, EventArgs e) //조회버튼
        {
            string Procedure = string.Empty;
            string Report = string.Empty;
            string DataSheet = string.Empty;

            if (rbl_view_type.SelectedValue.ToString()== "A") // 음악교실
            {
                Procedure = "dbo.USP_EDU_NEPESWAY_VIEW";
                Report = "rp_e_n_way_admin_music.rdlc";
                DataSheet = "DataSet1";
            }

            if (rbl_view_type.SelectedValue.ToString() == "B") //i훈련
            {
                Procedure = "dbo.USP_EDU_NEPESWAY_VIEW";
                Report = "rp_e_n_way_admin_book.rdlc";
                DataSheet = "DataSet1";
            }
           
            if (rbl_view_type.SelectedValue.ToString() == "C") // 마법노트
            {
                Procedure = "dbo.USP_EDU_NEPESWAY_VIEW";
                Report = "rp_e_n_way_admin_mabup.rdlc";
                DataSheet = "DataSet1";
            }

         

            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = Procedure;
            cmd.CommandTimeout = 0;
            

            SqlParameter param1 = new SqlParameter("@GUBUN", SqlDbType.VarChar, 2);
            SqlParameter param2 = new SqlParameter("@FR_YYYYMMDD", SqlDbType.VarChar, 8);
            SqlParameter param3 = new SqlParameter("@TO_YYYYMMDD", SqlDbType.VarChar, 8);
            SqlParameter param4 = new SqlParameter("@BUSOR_1 ", SqlDbType.VarChar, 30);
            SqlParameter param5 = new SqlParameter("@BUSOR_2 ", SqlDbType.VarChar, 30);
            SqlParameter param6 = new SqlParameter("@BUSOR_3 ", SqlDbType.VarChar, 30);
            SqlParameter param7 = new SqlParameter("@BUSOR_4 ", SqlDbType.VarChar, 30);
            SqlParameter param8 = new SqlParameter("@NAME", SqlDbType.VarChar, 10);
            SqlParameter param9 = new SqlParameter("@SNO ", SqlDbType.VarChar, 10);
            SqlParameter param10 = new SqlParameter("@GUNMU ", SqlDbType.VarChar, 20);

            
            param1.Value = rbl_view_type.SelectedValue.ToString();
            param2.Value = fr_yyyymmdd.Text;

            if (fr_yyyymmdd == null || fr_yyyymmdd.Text.Equals(""))
            {
                MessageBox.ShowMessage("시작일을 입력하세요.", this.Page);

                return;
            }
            param3.Value = to_yyyymmdd.Text;
            if (to_yyyymmdd == null || to_yyyymmdd.Text.Equals(""))
            {
                MessageBox.ShowMessage("종료일을 입력하세요.", this.Page);

                return;
            }
            param4.Value = tb_busor1.Text;
            param5.Value = tb_busor2.Text;
            param6.Value = tb_busor3.Text;
            param7.Value = tb_busor4.Text;
            param8.Value = tb_name.Text;
            param9.Value = tb_sno.Text;
            param10.Value = ddl_gunmu.SelectedValue.ToString();
           

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);
            cmd.Parameters.Add(param5);
            cmd.Parameters.Add(param6);
            cmd.Parameters.Add(param7);
            cmd.Parameters.Add(param8);
            cmd.Parameters.Add(param9);
            cmd.Parameters.Add(param10);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath(Report);
                ReportViewer1.LocalReport.DisplayName = "Nepesway실적조회" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = DataSheet;
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


