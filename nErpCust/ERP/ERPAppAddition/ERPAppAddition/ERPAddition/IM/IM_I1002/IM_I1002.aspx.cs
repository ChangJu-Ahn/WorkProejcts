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
using ERPAppAddition.ERPAddition.IM.IM_I1002;

namespace ERPAppAddition.ERPAddition.IM.IM_I1002
{
    public partial class IM_I1002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr = null;
        DataSet ds = new DataSet();
        cls_prod_qty_month cls_dbexe = new cls_prod_qty_month();
        //string ls_fr_dt, ls_to_dt;
        //int value;
        //string ls_report_nm, ls_sql, ls_ddl_sql, ls_cost_cd, ls_yyyymm, ls_cost_cd_gp, ls_item_cd_gp, ls_weight, sql, now_month, before_month;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                WebSiteCount();
                ReportViewer1.Reset();
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
            ReportViewer1.Reset();

            if (rbtn_view_type.SelectedValue == "ERP")
            {
                // 프로시져 실행: 기본데이타 생성
                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "dbo.usp_daily_stock_viewer";
                cmd.CommandTimeout = 30000;

                SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 6);
                SqlParameter param2 = new SqlParameter("@plant_cd", SqlDbType.VarChar, 10);
                SqlParameter param3 = new SqlParameter("@item_acct", SqlDbType.VarChar, 2);
                SqlParameter param4 = new SqlParameter("@item_cd", SqlDbType.VarChar, 20);
                SqlParameter param5 = new SqlParameter("@sl_cd", SqlDbType.VarChar, 20);
                SqlParameter param6 = new SqlParameter("@type", SqlDbType.VarChar, 20);
                param1.Value = tb_yyyymm.Text;
                if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null)
                    MessageBox.ShowMessage("조회년월을 선택해주세요", this.Page);
                param2.Value = ddl_plant_cd.SelectedValue;
                param3.Value = ddl_item_acct.SelectedValue;
                if (tb_item_cd.Text == "" || tb_item_cd.Text == null)
                    param4.Value = "%";
                else
                    param4.Value = tb_item_cd.Text;
                param5.Value = ddl_sl_cd.SelectedValue;
                param6.Value = rbtn_erp_view_type.SelectedValue;

                cmd.Parameters.Add(param1);
                cmd.Parameters.Add(param2);
                cmd.Parameters.Add(param3);
                cmd.Parameters.Add(param4);
                cmd.Parameters.Add(param5);
                cmd.Parameters.Add(param6);
            }
            if (rbtn_view_type.SelectedValue == "MES")
            {
                // 프로시져 실행: 기본데이타 생성
                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "dbo.usp_daily_stock_mes_qty";
                cmd.CommandTimeout = 30000;

                SqlParameter param1 = new SqlParameter("@yyyymm", SqlDbType.VarChar, 6);
                SqlParameter param2 = new SqlParameter("@plant_cd", SqlDbType.VarChar, 10);
                SqlParameter param3 = new SqlParameter("@item_nm", SqlDbType.VarChar, 50);
                SqlParameter param4 = new SqlParameter("@type", SqlDbType.VarChar, 20);
                param1.Value = tb_yyyymm.Text;
                if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null)
                    MessageBox.ShowMessage("조회년월을 선택해주세요", this.Page);
                param2.Value = ddl_plant_cd.SelectedValue;

                if (tb_item_nm.Text == "" || tb_item_nm.Text == null)
                    param3.Value = "%";
                else
                    param3.Value = tb_item_nm.Text;
                param4.Value = rbtn_mes_view_type.SelectedValue;

                cmd.Parameters.Add(param1);
                cmd.Parameters.Add(param2);
                cmd.Parameters.Add(param3);
                cmd.Parameters.Add(param4);
            }

            try
            {

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                string report_nm="";
                
                if (rbtn_view_type.SelectedValue == "ERP")
                {
                    if (rbtn_erp_view_type.SelectedValue == "A")
                        report_nm = "rp_im_i1002_erp_sum.rdlc"; //erp 집계조회
                    else if (rbtn_erp_view_type.SelectedValue == "B") /*상세조회-가로*/
                        report_nm = "rp_im_i1002_erp_detail_1.rdlc";
                    else
                        report_nm = "rp_im_i1002_erp_detail_2.rdlc";/*상세조회-세로*/
                }
                if (rbtn_view_type.SelectedValue == "MES")
                {
                    if (rbtn_mes_view_type.SelectedValue == "A")
                        report_nm = "rp_im_i1002_mes_sum.rdlc"; //mes 집계조회
                    else if (rbtn_mes_view_type.SelectedValue == "B")
                        report_nm = "rp_im_i1002_mes_detail.rdlc";//mes 상세조회         
                }
                ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                if (rbtn_view_type.SelectedValue =="ERP")
                   ReportViewer1.LocalReport.DisplayName = "일일재고금액 " + ddl_plant_cd.SelectedItem + DateTime.Now.ToShortDateString();
                if (rbtn_view_type.SelectedValue == "MES")
                    ReportViewer1.LocalReport.DisplayName = "MES공정실적 " + ddl_plant_cd.SelectedItem + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

                UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                Console.WriteLine("{0} Second exception caught.", ex);
            }
        }
        protected void bt_item_cd_Click(object sender, EventArgs e)
        {
            pop_gridview1.DataSourceID = "";
            pop_gridview1.DataSource = null;
            pop_gridview1.DataBind();
        }

        //팝업창에서 조회버튼 
        protected void bt_retrive_Click(object sender, EventArgs e)
        {
            pop_gridview1.DataSource = "";
            pop_gridview1.DataSource = SqlDataSource_pop_query;
            pop_gridview1.DataBind();
            pop_gridview1.Visible = true;
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Show();
        }
        // 팝업창-취소버튼클릭시 
        protected void bt_cancel_Click(object sender, EventArgs e)
        {
            //기존 보여졌던 데이타들을 안보이게 초기화
            pop_gridview1.DataSource = dr;
            pop_gridview1.DataBind();
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Hide();
        }
        protected void pop_gridview1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            //pageallow를 계속적으로 수행하기 위해 아래 코드가 필요
            pop_gridview1.PageIndex = e.NewPageIndex;
            pop_gridview1.DataBind();
            //새페이지를 눌렀을경우 gridview가 사라지기에 다시 조회하도록 조회버튼 호출
            bt_retrive_Click(this, e);
        }

        protected void pop_gridview1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //GridViewRow row = pop_gridview1.SelectedRow;
            //ls_item_cd = row.Cells[1].Text;
            //ls_item_nm = row.Cells[2].Text;


            //btn_pop_ok.Enabled = true;
        }
        // ok 버튼을 클릭하면 부모창에 값을 전달한다.
        protected void btn_pop_ok_Click(object sender, EventArgs e)
        {

            int i_chk_rowcnt = pop_gridview1.Rows.Count;
            string ls_chk_selectrowindex = pop_gridview1.SelectedIndex.ToString();

            if (ls_chk_selectrowindex != "-1")
            {
                GridViewRow row = pop_gridview1.SelectedRow;
                tb_item_cd.Text = row.Cells[1].Text;
                tb_item_nm.Text = row.Cells[2].Text;
            }
            pop_gridview1.DataSource = dr;
            pop_gridview1.DataBind();

        }

        
        protected void rbtn_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbtn_view_type.SelectedValue == "ERP")
            {
                rbtn_mes_view_type.Visible = false;
                rbtn_erp_view_type.Visible = true;
                ddl_item_acct.Enabled = true;
                ddl_sl_cd.Enabled = true;
                tb_item_cd.Enabled = true;
                bt_item_cd.Enabled = true;
                tb_item_nm.Enabled = true;
            }
            if (rbtn_view_type.SelectedValue == "MES")
            {
                rbtn_erp_view_type.Visible = false;
                rbtn_mes_view_type.Visible = true;
                ddl_item_acct.Enabled = false;
                ddl_sl_cd.Enabled = false;
                tb_item_cd.Enabled = true;
                bt_item_cd.Enabled = true;
                tb_item_nm.Enabled = true;

            }
        }
    }
}