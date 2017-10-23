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
using ERPAppAddition.ERPAddition.CM.CM_C3001;

namespace ERPAppAddition.ERPAddition.CM.CM_C3001
{
    public partial class CM_C3001 : System.Web.UI.Page
    {
        SqlConnection conn ; //= new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        DataSet ds = new DataSet();
        cls_prod_qty_month cls_dbexe = new cls_prod_qty_month();

        string strcon;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
               // chk_server();
                ReportViewer1.Reset();
            }
        }

        private void chk_server()
        {
            
            
        }

        protected void btn_request_Click(object sender, EventArgs e)
        {
            if ((tb_fr_dt.Text == "") || (tb_fr_dt.Text == null))
            {
                string script = "alert(\"조회년월을 선택하여 주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                tb_fr_dt.Focus();

            }
            else
            {
                ReportViewer1.Reset();

                // db선택
                if (rbl_server.SelectedValue == "ERP")
                    strcon = "server=192.168.10.15;database=nepes;uid=sa;pwd=nepes01!;";
                if (rbl_server.SelectedValue == "COST")
                    strcon = "server=192.168.10.98;database=nepes_cost;uid=sa;pwd=nepes123;";
                conn = new SqlConnection(strcon);

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "dbo.usp_view_mes_vs_erp_inv_qty";
                cmd.CommandTimeout = 30000;

                SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 6);
                SqlParameter param2 = new SqlParameter("@plant_cd", SqlDbType.VarChar, 10);
                SqlParameter param3 = new SqlParameter("@view_type", SqlDbType.VarChar, 10);

                param1.Value = tb_fr_dt.Text;
                param2.Value = ddl_plant_cd.SelectedValue;
                param3.Value = RadioButtonList1.SelectedValue;

                cmd.Parameters.Add(param1);
                cmd.Parameters.Add(param2);
                cmd.Parameters.Add(param3);                

                try
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    //ds_cm_c3001 ds = new ds_cm_c3001();
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    string report_nm = "";

                    if (RadioButtonList1.SelectedValue == "view1")
                        report_nm = "rp_cm_c3001.rdlc"; //erp 집계조회
                    else if (RadioButtonList1.SelectedValue == "view2") /*상세조회-가로*/
                        report_nm = "rp_cm_c3001_view2.rdlc";
                    else
                        report_nm = "rp_cm_c3001_view3.rdlc";/*상세조회-세로*/

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                    ReportViewer1.LocalReport.DisplayName = "MES VS ERP 재공 " + ddl_plant_cd.SelectedItem + DateTime.Now.ToShortDateString();
                    
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

        protected void rbl_server_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }

        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }
    }
}