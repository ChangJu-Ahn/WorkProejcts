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
using ERPAppAddition.ERPAddition.AM.AM_A7001;

namespace ERPAppAddition.ERPAddition.AM.AM_A7001
{
    public partial class AM_A7001 : System.Web.UI.Page
    {
        //SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlConnection conn;

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string ls_gl_no;
        string userid, db_name;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
            {
                db_name = Request.QueryString["db"].ToString();
                if (db_name.Length > 0)
                {
                    conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);
                }
                userid = Request.QueryString["userid"];

                Session["DBNM"] = Request.QueryString["db"].ToString();
                Session["User"] = Request.QueryString["userid"];
            }
            else
            {
                string script = "alert(\"프로그램 호출이 잘못되었습니다. 관리자에게 연락해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }

            Label1.Text = "*" + BIZ_AREA.SelectedItem.Text + " : ";
            Label2.Text = "*" + BIZ_AREA.SelectedItem.Text + "명 : ";

            InitDropDownList();
            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_" + dl_plant_cd.Text.Trim() + "BOM_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                _reportViewer.LocalReport.DataSources.Add(rds);

                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        //팝업창에서 조회버튼 
        protected void bt_retrive_Click(object sender, EventArgs e)
        {

            pop_gridview1.DataSource = "";
            if (BIZ_AREA.SelectedItem.Text == "주문처")
            {
                SqlDataSource3.SelectCommand = "SELECT DISTINCT TMP.BP_CD, TMP.BP_NM FROM (SELECT DISTINCT A.BP_CD, A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AR B WHERE A.BP_CD = B.DEAL_BP_CD AND B.CONF_FG = 'C' AND AR_DT <= '" + tb_fr_dt.Text + "') TMP WHERE TMP.BP_CD LIKE '" + tb_pop_item_cd.Text + "%' AND TMP.BP_NM LIKE '" + tb_pop_item_nm.Text + "%' ORDER BY TMP.BP_CD";
            }
            else
            {
                SqlDataSource3.SelectCommand = "SELECT DISTINCT TMP.BP_CD, TMP.BP_NM FROM (SELECT DISTINCT A.BP_CD, A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AR B WHERE A.BP_CD = B.PAY_BP_CD AND B.CONF_FG = 'C' AND AR_DT <= '" + tb_fr_dt.Text + "') TMP WHERE TMP.BP_CD LIKE '" + tb_pop_item_cd.Text + "%' AND TMP.BP_NM LIKE '" + tb_pop_item_nm.Text + "%' ORDER BY TMP.BP_CD";
            }

            SqlDataSource3.ConnectionString = ConfigurationManager.ConnectionStrings[db_name].ConnectionString;

            pop_gridview1.DataSource = SqlDataSource3;
            pop_gridview1.DataBind();
            pop_gridview1.Visible = true;
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Show();


        }
        // 부모창 - 조회버튼
        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            string sql;

            if (BIZ_AREA.SelectedItem.Text == "주문처")
            {
                sql = "USP_A_BOND_STATEMENT_DEAL_BP_SERCH '" + tb_fr_dt.Text.Trim() + "','" + tb_item_cd.Text.Trim() + "', '" + dl_plant_cd.Text.Trim() + "'";
                DataSet_AM_A7001 dt1 = new DataSet_AM_A7001();
                ReportViewer1.Reset();
                ReportCreator(dt1, sql, ReportViewer1, "AM_A7001_1.rdlc", "DataSet1");
            }
            else
            {
                sql = "EXEC USP_A_BOND_STATEMENT_PAY_BP_CD_SERCH '" + tb_fr_dt.Text.Trim() + "','" + tb_item_cd.Text.Trim() + "', '" + dl_plant_cd.Text.Trim() + "'";
                DataSet_AM_A7001 dt1 = new DataSet_AM_A7001();
                ReportViewer1.Reset();
                ReportCreator(dt1, sql, ReportViewer1, "AM_A7001_2.rdlc", "DataSet1");
            }
            
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
        // 팝업으로 품목코드 선택창 띄우기
        protected void bt_item_cd_Click(object sender, EventArgs e)
        {
            pop_gridview1.DataSourceID = "";
            pop_gridview1.DataSource = null;
            pop_gridview1.DataBind();
            
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

        /// <summary>
        /// BIZ_AREA_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void InitDropDownList()
        {
            //dl_plant_cd
            dl_plant_cd.DataSource = "";
            SqlDataSource1.SelectCommand = "SELECT DISTINCT BIZ_AREA_CD, BIZ_AREA_NM FROM B_BIZ_AREA ORDER BY BIZ_AREA_CD";
            SqlDataSource1.ConnectionString = ConfigurationManager.ConnectionStrings[db_name].ConnectionString;
            dl_plant_cd.DataSource = SqlDataSource1;
            dl_plant_cd.DataValueField = "BIZ_AREA_CD";
            dl_plant_cd.DataTextField = "BIZ_AREA_NM";
            dl_plant_cd.DataBind();
        }
    }
}