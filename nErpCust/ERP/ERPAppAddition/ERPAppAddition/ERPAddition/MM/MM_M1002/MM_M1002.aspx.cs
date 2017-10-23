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
using ERPAppAddition.ERPAddition.MM.MM_M1002;

namespace ERPAppAddition.ERPAddition.MM.MM_M1002
{
    public partial class MM_M1002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        
        string ls_item_cd = "", ls_item_nm = "";
        string ls_sql,ls_yyyymm, ls_sl_cd, ls_trsl_cd;
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                TimeSpan ts = new TimeSpan(-7, 0, 0, 0);

                DateTime date = DateTime.Now.Date.Add(ts);

                str_fr_dt.Text = date.Year.ToString("0000") + date.Month.ToString("00") + date.Day.ToString("00");
                str_to_dt.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");

                tb_item_cd.Text = "";
                tb_item_nm.Text = "";

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

        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            string sReportView = "rp_mm_m1002.rdlc";
            if(dl_plant_cd.SelectedIndex < 0 )
            {
                MessageBoxClass.ShowMessage("공장을 선택해 주세요", this.Page);
                return;
            }
            else
            {
                DataTable dtSL = GetSL_CD();
                if(dtSL.Rows.Count < 1)
                {
                    MessageBoxClass.ShowMessage("창고 코드에 문제가 있습니다.", this.Page);
                    return;
                }
                else
                {
                    ls_trsl_cd = dtSL.Rows[0]["TRNS_SL_CD"].ToString();
                    ls_sl_cd = dtSL.Rows[0]["SL_CD"].ToString();
                }

            }

            ls_item_cd = tb_item_cd.Text.Trim();
            //if (ls_item_cd.Length < 1 || ls_item_cd == "") ls_item_cd = "%";


            StringBuilder sSql = new StringBuilder();

            if (RadioButtonList1.SelectedValue == "view1")
            {    
                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	A.ITEM_CD");
                sSql.AppendLine("	, C.ITEM_NM");
                sSql.AppendLine("	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102) DT");
                sSql.AppendLine("	, SUM(A.QTY) AS QTY");
                sSql.AppendLine(" FROM I_GOODS_MOVEMENT_DETAIL A WITH(NOLOCK)");
                sSql.AppendLine("			INNER JOIN I_GOODS_MOVEMENT_HEADER B WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO");
                sSql.AppendLine("			INNER JOIN B_ITEM C WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_CD = C.ITEM_CD");
                sSql.AppendLine(" WHERE A.PLANT_CD = '" + dl_plant_cd.Text + "' ");
                sSql.AppendLine("  AND B.DOCUMENT_DT BETWEEN '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "'");
                if (!(ls_item_cd.Length < 1 || ls_item_cd == ""))
                {
                    sSql.AppendLine("  AND A.ITEM_CD = '" + ls_item_cd + "'");
                }
                sSql.AppendLine("  AND A.DELETE_FLAG <> 'Y'");
                sSql.AppendLine("  AND A.TRNS_TYPE = 'ST'");
                sSql.AppendLine("  AND A.MOV_TYPE IN ('I04', 'TX7')");
                //sSql.AppendLine("  AND A.TRNS_SL_CD = '" + ls_trsl_cd + "'");
                sSql.AppendLine("  AND A.SL_CD = '" + ls_sl_cd + "'");
                sSql.AppendLine(" GROUP BY A.ITEM_CD");
                sSql.AppendLine(" 	, C.ITEM_NM");
                sSql.AppendLine(" 	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102)");
                sSql.AppendLine(" ORDER BY A.ITEM_CD, DT");

            }

            else if (RadioButtonList1.SelectedValue == "view2")
            {

                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	A.ITEM_CD");
                sSql.AppendLine("	, C.ITEM_NM");
                sSql.AppendLine("	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102) DT");
                sSql.AppendLine("	, SUM(A.QTY) AS QTY");
                sSql.AppendLine(" FROM I_GOODS_MOVEMENT_DETAIL A WITH(NOLOCK)");
                sSql.AppendLine("			INNER JOIN I_GOODS_MOVEMENT_HEADER B WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO");
                sSql.AppendLine("			INNER JOIN B_ITEM C WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_CD = C.ITEM_CD");
                sSql.AppendLine(" WHERE A.PLANT_CD = '" + dl_plant_cd.Text + "' ");
                sSql.AppendLine("  AND B.DOCUMENT_DT BETWEEN '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "'");
                if (!(ls_item_cd.Length < 1 || ls_item_cd == ""))
                {
                    sSql.AppendLine("  AND A.ITEM_CD = '" + ls_item_cd + "'");
                }
                sSql.AppendLine("  AND A.DELETE_FLAG <> 'Y'");
                sSql.AppendLine("  AND A.TRNS_TYPE = 'OI'");
                sSql.AppendLine("  AND A.MOV_TYPE = 'I03'");
                sSql.AppendLine("  AND A.SL_CD = '" + ls_sl_cd + "'");
                sSql.AppendLine(" GROUP BY A.ITEM_CD");
                sSql.AppendLine(" 	, C.ITEM_NM");
                sSql.AppendLine(" 	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102)");
                sSql.AppendLine(" ORDER BY A.ITEM_CD, DT");
            }

            else if (RadioButtonList1.SelectedValue == "InOut")
            {


                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	A.ITEM_CD");
                sSql.AppendLine("	, A.ITEM_NM");
                sSql.AppendLine("	, A.BASE_UNIT");
                sSql.AppendLine("	, A.DT");
                sSql.AppendLine("	, MAX(A.STOCK_QTY) AS STOCK_QTY");
                sSql.AppendLine("	, SUM(A.IN_QTY) AS IN_QTY");
                sSql.AppendLine("	, SUM(A.OUT_QTY) AS OUT_QTY");
                sSql.AppendLine(" FROM (");
                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	A.ITEM_CD");
                sSql.AppendLine("	, C.ITEM_NM");
                sSql.AppendLine("	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102) DT");
                sSql.AppendLine("	, NULL AS STOCK_QTY");
                sSql.AppendLine("	, SUM(A.QTY) AS IN_QTY");
                sSql.AppendLine("	, NULL AS OUT_QTY");
                sSql.AppendLine(" FROM I_GOODS_MOVEMENT_DETAIL A WITH(NOLOCK)");
                sSql.AppendLine("			INNER JOIN I_GOODS_MOVEMENT_HEADER B WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO");
                sSql.AppendLine("			INNER JOIN B_ITEM C WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_CD = C.ITEM_CD");
                sSql.AppendLine(" WHERE A.PLANT_CD = '" + dl_plant_cd.Text + "' ");
                sSql.AppendLine("  AND B.DOCUMENT_DT BETWEEN '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "'");
                if (!(ls_item_cd.Length < 1 || ls_item_cd == ""))
                {
                    sSql.AppendLine("  AND A.ITEM_CD = '" + ls_item_cd + "'");
                }
                sSql.AppendLine("  AND A.DELETE_FLAG <> 'Y'");
                sSql.AppendLine("  AND A.TRNS_TYPE = 'ST'");
                sSql.AppendLine("  AND A.MOV_TYPE IN ('I04', 'TX7')");
                //sSql.AppendLine("  AND A.TRNS_SL_CD = '" + ls_trsl_cd + "'");
                sSql.AppendLine("  AND A.SL_CD = '" + ls_sl_cd + "'");
                sSql.AppendLine(" GROUP BY A.ITEM_CD");
                sSql.AppendLine(" 	, C.ITEM_NM");
                sSql.AppendLine(" 	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102)");

                sSql.AppendLine(" UNION ALL ");

                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	A.ITEM_CD");
                sSql.AppendLine("	, C.ITEM_NM");
                sSql.AppendLine("	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102) DT");
                sSql.AppendLine("	, NULL AS STOCK_QTY");
                sSql.AppendLine("	, NULL AS IN_QTY");
                sSql.AppendLine("	, SUM(A.QTY) AS OUT_QTY");
                sSql.AppendLine(" FROM I_GOODS_MOVEMENT_DETAIL A WITH(NOLOCK)");
                sSql.AppendLine("			INNER JOIN I_GOODS_MOVEMENT_HEADER B WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO");
                sSql.AppendLine("			INNER JOIN B_ITEM C WITH(NOLOCK)");
                sSql.AppendLine("			ON A.ITEM_CD = C.ITEM_CD");
                sSql.AppendLine(" WHERE A.PLANT_CD = '" + dl_plant_cd.Text + "' ");
                sSql.AppendLine("  AND B.DOCUMENT_DT BETWEEN '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "'");
                if (!(ls_item_cd.Length < 1 || ls_item_cd == ""))
                {
                    sSql.AppendLine("  AND A.ITEM_CD = '" + ls_item_cd + "'");
                }
                sSql.AppendLine("  AND A.DELETE_FLAG <> 'Y'");
                sSql.AppendLine("  AND A.TRNS_TYPE = 'OI'");
                sSql.AppendLine("  AND A.MOV_TYPE = 'I03'");
                sSql.AppendLine("  AND A.SL_CD = '" + ls_sl_cd + "'");
                sSql.AppendLine(" GROUP BY A.ITEM_CD");
                sSql.AppendLine(" 	, C.ITEM_NM");
                sSql.AppendLine(" 	, A.BASE_UNIT");
                sSql.AppendLine("	, CONVERT(VARCHAR(10), B.DOCUMENT_DT, 102)");

                sSql.AppendLine(" UNION ALL ");

                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("	B.ITEM_CD");
                sSql.AppendLine("	, B.ITEM_NM");
                sSql.AppendLine("	, B.BASIC_UNIT");
                sSql.AppendLine("	, 'AA' DT");
                sSql.AppendLine("	, A.GOOD_ON_HAND_QTY + A.BAD_ON_HAND_QTY+A.STK_ON_INSP_QTY+A.STK_IN_TRNS_QTY AS STOCK_QTY");
                sSql.AppendLine("	, NULL AS IN_QTY");
                sSql.AppendLine("	, NULL AS OUT_QTY");
                sSql.AppendLine(" FROM I_ONHAND_STOCK A WITH(NOLOCK)");
                sSql.AppendLine("	INNER JOIN B_ITEM_BY_PLANT E(NOLOCK) ON A.PLANT_CD = E.PLANT_CD");
                sSql.AppendLine("		   AND A.ITEM_CD = E.ITEM_CD");
                sSql.AppendLine("	INNER JOIN B_STORAGE_LOCATION C(NOLOCK) ON A.SL_CD = C.SL_CD");
                sSql.AppendLine("	INNER JOIN B_PLANT D(NOLOCK) ON A.PLANT_CD = D.PLANT_CD");
                sSql.AppendLine("	INNER JOIN B_ITEM B(NOLOCK) ON A.ITEM_CD = B.ITEM_CD");

                sSql.AppendLine(" WHERE D.PLANT_CD  = '" + dl_plant_cd.Text + "' ");
                sSql.AppendLine("  AND A.GOOD_ON_HAND_QTY <> 0");
                if (!(ls_item_cd.Length < 1 || ls_item_cd == ""))
                {
                    sSql.AppendLine("  AND A.ITEM_CD = '" + ls_item_cd + "'");
                }
                sSql.AppendLine("  AND C.SL_CD = '" + ls_sl_cd + "'");
                
                sSql.AppendLine(")A");
                sSql.AppendLine(" GROUP BY A.ITEM_CD, A.ITEM_NM, A.BASE_UNIT, A.DT");

                sSql.AppendLine(" ORDER BY A.ITEM_CD, A.DT");

                sReportView = "rp_mm_m1002_All.rdlc";
            }

            
            ds_mm_m1002 dt1 = new ds_mm_m1002();
            ReportViewer1.Reset();
            ReportCreator(dt1, sSql.ToString(), ReportViewer1, sReportView, "DataSet1");
        
        }

        protected void bt_item_cd_Click(object sender, EventArgs e)
        {
            //pop_gridview1.DataSourceID = "";
            //pop_gridview1.DataSource = null;
            //pop_gridview1.DataBind();

            tb_pop_item_cd.Text = tb_item_cd.Text;
            tb_pop_item_nm.Text = tb_item_nm.Text;

            DataSourceSelectArguments arg = new DataSourceSelectArguments();

            SqlDataSource2.Select(arg);
            
            SearchPopUp();
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

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                //dr = cmd.ExecuteReader();
                //ds.Tables[0].Load(dr);
                //dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_" + dl_plant_cd.Text.Trim() + "_" + str_fr_dt.Text + "_" + str_to_dt.Text + "_" + RadioButtonList1.SelectedItem.Text + "_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = dt;
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

        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {
          
            //tb_item_cd.Text="";
            //tb_item_nm.Text = "";

            ReportViewer1.Reset();

        
        }
        //팝업창에서 조회버튼 
        protected void bt_retrive_Click(object sender, EventArgs e)
        {
            SearchPopUp();
        }

        private void SearchPopUp()
        {
            pop_gridview1.DataSource = "";
            pop_gridview1.DataSource = SqlDataSource2;
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
            //bt_retrive_Click(this, e);
            SearchPopUp();
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

        private DataTable GetSL_CD()
        {
            StringBuilder sSql = new StringBuilder();
            DataTable dt = new DataTable();
            if (dl_plant_cd.SelectedIndex > -1)
            {
                sSql.AppendLine(" SELECT ");
                sSql.AppendLine("  MAX(SL_CD) AS SL_CD");
                sSql.AppendLine(" , MAX(TRNS_SL) AS TRNS_SL_CD");
                sSql.AppendLine(" FROM");
                sSql.AppendLine(" ( SELECT MAX(SL_CD) AS SL_CD");
                sSql.AppendLine("	,'' AS TRNS_SL");
                sSql.AppendLine(" FROM B_STORAGE_LOCATION WITH (NOLOCK)");
                sSql.AppendLine("	WHERE 1 = 1 ");
                sSql.AppendLine("   AND PLANT_CD = '" + dl_plant_cd.Text + "'");
                sSql.AppendLine("  	AND SL_CD LIKE '%6000'");
                sSql.AppendLine("	AND SL_NM LIKE '%소모품%'");
                sSql.AppendLine(" UNION ");
                sSql.AppendLine("  SELECT '' AS SL_CD");
                sSql.AppendLine("	, MAX(SL_CD) AS TRNS_SL_CD");
                sSql.AppendLine(" FROM B_STORAGE_LOCATION WITH (NOLOCK)");
                sSql.AppendLine("	WHERE 1 = 1 ");
                sSql.AppendLine("   AND PLANT_CD = '" + dl_plant_cd.Text + "'");
                sSql.AppendLine("  	AND SL_CD LIKE '%3000'");
                sSql.AppendLine("	AND SL_NM LIKE '%자재%'");
                sSql.AppendLine(" )A ");

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;

                try
                {
                    cmd.CommandText = sSql.ToString();

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(dt);

                    //dr = cmd.ExecuteReader();
                    //ds.Tables[0].Load(dr);
                    //dr.Close();

                    ls_sl_cd = dt.Rows[0]["SL_CD"].ToString();
                    ls_trsl_cd = dt.Rows[0]["TRNS_SL"].ToString();

                }
                catch
                {
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
            }
            return dt;
        }

        protected void dl_plant_cd_SelectedIndexChanged(object sender, EventArgs e)
        {

            tb_item_cd.Text="";
            tb_item_nm.Text = "";

            ReportViewer1.Reset();
        }

        protected void ModalPopupExtender1_Load(object sender, EventArgs e)
        {

        }

        protected void ModalPopupExtender1_PreRender(object sender, EventArgs e)
        {

        }
    }
}