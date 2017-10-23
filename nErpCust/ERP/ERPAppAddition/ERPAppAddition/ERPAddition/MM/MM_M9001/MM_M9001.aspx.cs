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
using ERPAppAddition.ERPAddition.MM.MM_M9001;


namespace ERPAppAddition.ERPAddition.MM.MM_M9001
{
    public partial class MM_M9001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            /*
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
                //string script = "alert(\"프로그램 호출이 잘못되었습니다. 관리자에게 연락해주세요.\");";
                //ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                db_name = "nepes";
                if (db_name.Length > 0)
                {
                    conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);
                }
            }
            */

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
            pop_gridview1.DataSource = SqlDataSource2;
            pop_gridview1.DataBind();
            pop_gridview1.Visible = true;
            pop_gridview1.SelectedIndex = -1;
            ModalPopupExtender1.Show();
            
            
        }
        // 부모창 - 조회버튼
        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            string item_cd = string.Empty;
            string fr_dt = string.Empty;
            string to_dt = string.Empty;

            if (!string.IsNullOrEmpty(str_fr_dt.Text.Trim()))
            {
                fr_dt = str_fr_dt.Text.Trim() + "070000";
            }
            if (!string.IsNullOrEmpty(str_to_dt.Text.Trim()))
            {
                to_dt = DateTime.ParseExact(str_to_dt.Text.Substring(0, 8).ToLower(), "yyyyMMdd", null).AddMonths(0).AddDays(1).ToString("yyyyMMdd") + "070000";
            }

            item_cd = tb_item_cd.Text.Trim();
            if (item_cd.Length < 1 || item_cd == "") item_cd = "%";

            //string sql = "select * from ufn_p1406ma_nepes('" + dl_plant_cd.Text.Trim() + "', '" + item_cd + "'  ) " +
            //             " order by B_ITEM_CD, ITEM_ACCT, LVL, CONVERT(INT,c_seq) ";
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT B.ITEM_SEQ, ");
            sbSql.AppendLine(" 	   B.ITEM_CD, ");
            sbSql.AppendLine(" 	   (SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = B.ITEM_CD) ITEM_NM, ");
            sbSql.AppendLine(" 	   (SELECT BASIC_UNIT FROM B_ITEM WHERE ITEM_CD = B.ITEM_CD) BASIC_UNIT, ");
            sbSql.AppendLine(" 	   B.REQ_QTY, ");
            sbSql.AppendLine(" 	   B.ISSUE_QTY, ");
            sbSql.AppendLine(" 	   B.PRODT_ORDER_NO, ");
            sbSql.AppendLine(" 	   B.REMARK REMARK1, ");
            sbSql.AppendLine(" 	   A.ISSUE_REQ_NO, ");
            sbSql.AppendLine(" 	   A.TRNS_TYPE, ");
            sbSql.AppendLine(" 	   CONVERT(CHAR(10), A.REQ_DT, 120) REQ_DT, ");
            sbSql.AppendLine(" 	   A.ISSUE_TYPE, ");
            sbSql.AppendLine(" 	   A.MOV_TYPE, ");
            sbSql.AppendLine(" 	   A.DEPT_CD, ");
            sbSql.AppendLine(" 	   A.EMP_NO, ");
            sbSql.AppendLine(" 	   A.REMARK, ");
            sbSql.AppendLine(" 	   A.CONFIRM_FLAG, ");
            sbSql.AppendLine("        [DBO].[UFN_GETMOVTYPE](TRNS_TYPE, MOV_TYPE ) POTYPECDNM, ");
            sbSql.AppendLine("        (SELECT NAME FROM HAA010T WHERE EMP_NO = A.EMP_NO) EMP_NM, ");
            sbSql.AppendLine("        DBO.UFN_GETDEPTNAME(DBO.UFN_H_GET_DEPT_CD(A.EMP_NO, GETDATE()), GETDATE()) DEPT_NM, ");
            sbSql.AppendLine(" 	   CONVERT(CHAR(10), B.LIMIT_DT, 120) LIMIT_DT, ");
            sbSql.AppendLine(" 	   A.LOC ");
            sbSql.AppendLine("   FROM M_ISSUE_REQ_HDR_KO441 A, M_ISSUE_REQ_DTL_KO441 B ");
            sbSql.AppendLine("  WHERE A.PLANT_CD = B.PLANT_CD ");
            sbSql.AppendLine("    AND A.ISSUE_REQ_NO = B.ISSUE_REQ_NO ");
            sbSql.AppendLine("    AND A.PLANT_CD = '" + dl_plant_cd.Text.Trim() + "' ");
            sbSql.AppendLine("    AND B.ITEM_CD LIKE '" + item_cd.Trim() + "' ");

            if (!string.IsNullOrEmpty(fr_dt.Trim()))
            {
                sbSql.AppendLine("    AND CONVERT(CHAR(10), A.REQ_DT, 120) >= '" + fr_dt.Substring(0, 8) + "' ");
            }
            if (!string.IsNullOrEmpty(to_dt.Trim()))
            {
                sbSql.AppendLine("    AND CONVERT(CHAR(10), A.REQ_DT, 120) <= '" + to_dt.Substring(0, 8) + "' ");
            }

            ds_mm_m9001 dt1 = new ds_mm_m9001();
            ReportViewer1.Reset();
            ReportCreator(dt1, sbSql.ToString(), ReportViewer1, "MM_M9001.rdlc", "DataSet1");
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
        
    }
}