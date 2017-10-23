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
using ERPAppAddition.ERPAddition.MM.MM_M1001;

namespace ERPAppAddition.ERPAddition.MM.MM_M1001
{
    public partial class MM_M1001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        
        string ls_item_cd = "", ls_item_nm = "";
        string ls_sql, ls_sql2, ls_sql3, ls_yyyymm, ls_sl_cd;
        protected void Page_Load(object sender, EventArgs e)
        {
            //미입고내역 조회시 창고 불필요
            if (RadioButtonList1.SelectedValue == "view2")
            {
                ddl_sl_cd.Visible = false;
            }
            else
            {
                ddl_sl_cd.Visible = true;
            }
            tb_item_cd.Text = "";
            tb_item_nm.Text = "";

            ReportViewer1.Reset();
            WebSiteCount();
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
            ls_item_cd = tb_item_cd.Text.Trim();
            ls_sl_cd = ddl_sl_cd.SelectedValue.ToString();
            if (ls_item_cd.Length < 1 || ls_item_cd == "") ls_item_cd = "%";

            //if (tb_yyyymm.Text.Length < 1 || tb_yyyymm.Text == "")
            //{
            //    MessageBoxClass.ShowMessage("조회년월을 입력해주세요", this.Page);
            //    return;
            //}


            //string ls_fr_dt, ls_to_dt;
            //ls_fr_dt = tb_yyyymm.Text + "01";
            //DateTime ld_fr_dt;
            //ld_fr_dt = DateTime.ParseExact(ls_fr_dt, "yyyyMMdd", null);
            //ls_to_dt = ld_fr_dt.AddMonths(1).AddDays(0 - ld_fr_dt.Day).ToString("yyyyMMdd");
            

            if (RadioButtonList1.SelectedValue == "view1")
            {
               // ls_sl_cd = ddl_sl_cd.SelectedValue.ToString();
                ls_sql = " SELECT A.ITEM_CD,B.ITEM_NM,T.MSDS,B.BASIC_UNIT,CONVERT(VARCHAR(10),A.MVMT_DT,102) DT,CASE WHEN F.RET_FLG='Y' THEN SUM(A.MVMT_QTY) * (-1) ELSE SUM(A.MVMT_QTY) END QTY " +
                         "  FROM M_PUR_GOODS_MVMT A    " +
                         "       INNER JOIN   B_ITEM B    ON A.ITEM_CD = B.ITEM_CD    " +
                         "       LEFT JOIN B_ITEM_DTL T   ON A.ITEM_CD = T.ITEM_CD    " +
                         "        INNER JOIN   M_MVMT_TYPE F    ON A.IO_TYPE_CD = F.IO_TYPE_CD AND F.RCPT_FLG <> 'N'    " +
                         "        LEFT OUTER JOIN   M_PUR_ORD_DTL G ON A.PO_NO = G.PO_NO AND A.PO_SEQ_NO = G.PO_SEQ_NO    " +
                         "        INNER JOIN   B_PUR_GRP H  ON A.PUR_GRP = H.PUR_GRP   " +
                         "  WHERE A.PLANT_CD like '" + dl_plant_cd.Text + "' " +
                         "    AND A.ITEM_CD like '" + ls_item_cd + "'  " +
                         "    AND A.MVMT_DT between '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "' " +
                         "    AND A.MVMT_SL_CD like '" + ls_sl_cd + "'   " +
                         " GROUP BY A.ITEM_CD,B.ITEM_NM,B.BASIC_UNIT,A.MVMT_DT, F.RET_FLG ,T.MSDS" +
                         " ORDER BY A.ITEM_CD";
            }

            if (RadioButtonList1.SelectedValue == "view2")
            {
                ls_sql = "SELECT B.ITEM_CD,F.ITEM_NM,T.MSDS,F.BASIC_UNIT,CONVERT(VARCHAR(10),B.DLVY_DT,102) DT,(SUM(B.PO_QTY) - SUM(B.RCPT_QTY)) QTY " +
                         "FROM M_PUR_ORD_HDR A, M_PUR_ORD_DTL B, B_BIZ_PARTNER C, B_PLANT D,  B_PUR_GRP E,  B_ITEM F, M_CONFIG_PROCESS G ,B_ITEM_DTL T   " +
                         "WHERE A.PO_NO = B.PO_NO AND A.BP_CD = C.BP_CD  " +
                         "AND B.PLANT_CD = D.PLANT_CD AND A.PUR_GRP *= E.PUR_GRP      " +
                         "AND B.ITEM_CD = F.ITEM_CD AND B.ITEM_CD *= T.ITEM_CD AND A.PO_TYPE_CD = G.PO_TYPE_CD AND (B.PO_QTY - B.RCPT_QTY) > 0  " +
                         "AND B.CLS_FLG = 'N'     " +
                         "AND B.PLANT_CD like '" + dl_plant_cd.Text + "' " +
                         "AND B.ITEM_CD like '" + ls_item_cd + "'  " +
                         "AND A.BP_CD >= '' AND A.BP_CD <= 'ZZZZZZZZZ'     " +
                         "AND B.DLVY_DT between '" + str_fr_dt.Text + "' AND '" + str_to_dt.Text + "' " ;
                if (view2_release_flg.Checked)
                    ls_sql2 = "AND A.RELEASE_FLG like '%' ";
                else
                    ls_sql2 = "AND A.RELEASE_FLG like 'Y' ";
                ls_sql3= "AND A.TRACKING_NO >= ''    " +
                         "AND ((B.PLANT_CD = 'P04' AND B.DLVY_DT >= '2008-09-01') OR (B.PLANT_CD <> 'P04' AND B.DLVY_DT >= '1900-01-01')) " +
                         "GROUP BY B.ITEM_CD,F.ITEM_NM,F.BASIC_UNIT,B.DLVY_DT ,T.MSDS " +
                         "Order By B.ITEM_CD ASC  ";
                ls_sql = ls_sql + ls_sql2 + ls_sql3;
            }
            if (RadioButtonList1.SelectedValue == "view3")
            {
                //ls_sl_cd = ddl_sl_cd.SelectedValue.ToString();
                ls_sql = " SELECT  B.ITEM_CD,       C.ITEM_NM,   T.MSDS,    B.BASE_UNIT BASIC_UNIT,       CONVERT(VARCHAR(10),A.DOCUMENT_DT ,102) DT,       sum(B.QTY) qty " +
                         " FROM I_GOODS_MOVEMENT_HEADER A inner join    (I_GOODS_MOVEMENT_DETAIL B " +
                         "	left join B_ITEM C on B.ITEM_CD = C.ITEM_CD     " +
                         "	left join B_ITEM_BY_PLANT D on B.PLANT_CD = D.PLANT_CD    AND B.ITEM_CD = D.ITEM_CD)    on A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO AND A.DOCUMENT_YEAR = B.DOCUMENT_YEAR  LEFT JOIN B_ITEM_DTL T ON B.ITEM_CD = T.ITEM_CD" +
                         " where B.DELETE_FLAG <> 'Y' and  B.PLANT_CD like '" + dl_plant_cd.Text + "' " +
                         " AND 	B.ITEM_CD LIKE  '" + ls_item_cd + "'  " +
                         " AND 	A.DOCUMENT_DT >= '" + str_fr_dt.Text + "' AND A.DOCUMENT_DT <= '" + str_to_dt.Text + "' " +
                         " AND 	B.SL_CD LIKE '" + ls_sl_cd + "'  and  B.MOV_TYPE =  'TX1'  and  B.TRNS_TYPE =  'ST'  " +
                         " group by B.ITEM_CD,       C.ITEM_NM,   A.DOCUMENT_DT,       B.BASE_UNIT  ,T.MSDS" +
                         " Order By B.ITEM_CD ASC , A.DOCUMENT_DT ASC  "; 
            }

            
            ds_mm_m1001 dt1 = new ds_mm_m1001();
            ReportViewer1.Reset();
            ReportCreator(dt1, ls_sql, ReportViewer1, "rp_mm_m1001.rdlc", "DataSet1");
        
        }

        protected void bt_item_cd_Click(object sender, EventArgs e)
        {
            pop_gridview1.DataSourceID = "";
            pop_gridview1.DataSource = null;
            pop_gridview1.DataBind();
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

                _reportViewer.LocalReport.DisplayName = "REPORT_" + dl_plant_cd.Text.Trim() + "_" + str_fr_dt.Text + "_" + str_to_dt.Text + "_" + RadioButtonList1.SelectedItem.Text + "_" + DateTime.Now.ToShortDateString();
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

        protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //미입고내역 조회시 창고 불필요
            if (RadioButtonList1.SelectedValue == "view2")
            {
                ddl_sl_cd.Visible = false;
                view2_release_flg.Visible = true;
            }
            else
            {
                ddl_sl_cd.Visible = true;
                view2_release_flg.Visible = false; //발주미확정포함 감추기
            }
            tb_item_cd.Text="";
            tb_item_nm.Text = "";

            ReportViewer1.Reset();

        
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
    }
}