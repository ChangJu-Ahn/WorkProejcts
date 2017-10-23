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
using ERPAppAddition.ERPAddition.MM.MM_M6001;

namespace ERPAppAddition.ERPAddition.MM.MM_M6001
{
    public partial class MM_M6001 : System.Web.UI.Page
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
            cmd.CommandText = "dbo.USP_IO_VIEW";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@PLANT_CD", SqlDbType.VarChar, 10);
            SqlParameter param2 = new SqlParameter("@ITEM_CD", SqlDbType.VarChar, 20);
            SqlParameter param3 = new SqlParameter("@BP_CD", SqlDbType.VarChar, 20);
            SqlParameter param4 = new SqlParameter("@MVMT_FR_DT", SqlDbType.VarChar, 20);
            SqlParameter param5 = new SqlParameter("@MVMT_TO_DT", SqlDbType.VarChar, 20);
            SqlParameter param6 = new SqlParameter("@SL_CD", SqlDbType.VarChar, 20);
            SqlParameter param7 = new SqlParameter("@IO_TYPE", SqlDbType.VarChar, 20);
            SqlParameter param8 = new SqlParameter("@FROM_INSRT_DT", SqlDbType.VarChar, 20);
            SqlParameter param9 = new SqlParameter("@TO_INSRT_DT", SqlDbType.VarChar, 20);

            string sql;
            string PLANT_CD, ITEM_CD, BP_CD, MVMT_FR_DT, MVMT_TO_DT, SL_CD, IO_TYPE;
            PLANT_CD = dl_plant_cd.SelectedValue;
            ITEM_CD = tb_item_cd.Text;
            BP_CD = tb_bp_cd.Text;
            MVMT_FR_DT = str_fr_dt.Text;
            MVMT_TO_DT = str_to_dt.Text;
            SL_CD = dl_sl_cd.SelectedValue;
            IO_TYPE = dl_io_type.SelectedValue;


            param1.Value = dl_plant_cd.SelectedValue;
            if (PLANT_CD == null || PLANT_CD == "")
                PLANT_CD = "%";

            param2.Value = tb_item_cd.Text;
            if (ITEM_CD == null || ITEM_CD == "")
                ITEM_CD = "%";

            param3.Value = tb_bp_cd.Text;
            if (BP_CD == null || BP_CD == "")
                BP_CD = "%";

            param4.Value = str_fr_dt.Text;
            if (MVMT_FR_DT == null || MVMT_FR_DT == "")
                MVMT_FR_DT = "19900101";
            param5.Value = str_to_dt.Text;
            if (MVMT_TO_DT == null || MVMT_TO_DT == "")
                MVMT_TO_DT = "29991231";

            param6.Value = dl_sl_cd.SelectedValue;
            if (SL_CD == null || SL_CD == "")
                SL_CD = "%";

            param7.Value = dl_io_type.SelectedValue;
            if (IO_TYPE == null || IO_TYPE == "")
                IO_TYPE = "%";

            param8.Value = insrt_Fr_dt.Text;
            if (IO_TYPE == null || IO_TYPE == "")
                IO_TYPE = "%";

            param9.Value = insrt_To_dt.Text;
            if (IO_TYPE == null || IO_TYPE == "")
                IO_TYPE = "%";

            cmd.Parameters.Add(param1);
            cmd.Parameters.Add(param2);
            cmd.Parameters.Add(param3);
            cmd.Parameters.Add(param4);
            cmd.Parameters.Add(param5);
            cmd.Parameters.Add(param6);
            cmd.Parameters.Add(param7);
            cmd.Parameters.Add(param8);
            cmd.Parameters.Add(param9);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_mm_m6001.rdlc");
                ReportViewer1.LocalReport.DisplayName = "입고상세조회(NEPES)" + DateTime.Now.ToShortDateString();
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
            }
        }
        protected void bt_item_cd_Click(object sender, EventArgs e) //품목 조회 버튼 클릭
        {
            pop_gridview1.DataSourceID = "";
            pop_gridview1.DataSource = null;
            pop_gridview1.DataBind();
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

        //거래처 조회 팝업
        protected void bt_bp_cd_Click1(object sender, EventArgs e) //거래처 조회 버튼 클릭
        {
            Response.Write("<script>window.open('pop_mm_m6001.aspx?pgid=mm_m6001&popupid=1','','top=100,left=100,width=1000,height=600')</script>");
        }

        //팝업창에서 조회버튼 
        //protected void btn_pop_retirve2_Click(object sender, EventArgs e)
        //{
        //    pop_gridview2.DataSource = "";
        //    pop_gridview2.DataSource = SqlDataSource4;
        //    pop_gridview2.DataBind();
        //    pop_gridview2.Visible = true;
        //    pop_gridview2.SelectedIndex = -1;
        //    bt_bp_cd_ModalPopupExtender.Show();
        //}
        //// 팝업창-취소버튼클릭시 
        //protected void btn_pop_cancel2_Click(object sender, EventArgs e)
        //{
        //    //기존 보여졌던 데이타들을 안보이게 초기화
        //    pop_gridview2.DataSource = dr;
        //    pop_gridview2.DataBind();
        //    pop_gridview2.SelectedIndex = -1;
        //    ModalPopupExtender1.Hide();
        //}
        //protected void pop_gridview2_PageIndexChanging(object sender, GridViewPageEventArgs e)
        //{
        //    //pageallow를 계속적으로 수행하기 위해 아래 코드가 필요
        //    pop_gridview2.PageIndex = e.NewPageIndex;
        //    pop_gridview2.DataBind();
        //    //새페이지를 눌렀을경우 gridview가 사라지기에 다시 조회하도록 조회버튼 호출
        //    bt_retrive_Click(this, e);
        //}

        //protected void pop_gridview2_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    //GridViewRow row = pop_gridview1.SelectedRow;
        //    //ls_item_cd = row.Cells[1].Text;
        //    //ls_item_nm = row.Cells[2].Text;


        //    //btn_pop_ok.Enabled = true;
        //}

        //// ok 버튼을 클릭하면 부모창에 값을 전달한다.
        //protected void btn_pop_ok2_Click(object sender, EventArgs e)
        //{

        //    int i_chk_rowcnt = pop_gridview2.Rows.Count;
        //    string ls_chk_selectrowindex = pop_gridview2.SelectedIndex.ToString();

        //    if (ls_chk_selectrowindex != "-1")
        //    {
        //        GridViewRow row = pop_gridview2.SelectedRow;
        //        tb_bp_cd.Text = row.Cells[1].Text;
        //        tb_bp_nm.Text = row.Cells[2].Text;
        //    }
        //    pop_gridview2.DataSource = dr;
        //    pop_gridview2.DataBind();

        //}

    }
}
