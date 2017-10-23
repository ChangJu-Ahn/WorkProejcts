using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
//using System.Data.OracleClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;
namespace ERPAppAddition.ERPAddition.SM.sm_s4001
{
    public partial class sm_s4001 : System.Web.UI.Page
    {
        //string strConn = ConfigurationManager.AppSettings["connectionKey"];

        string sql_cust_cd;

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleConnection conn_if = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_UNIERP"].ConnectionString);

        //SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_test1"].ConnectionString);

        SqlCommand cmd_erp = new SqlCommand();
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;
        SqlDataReader dr_erp;
        OracleDataAdapter sqlAdapter1;
        DataSet ds = new DataSet();

        DataSet ds1 = new DataSet();

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;

        int value, chk_save_yn = 0;
        string userid, db_name;


        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
                Label20.Visible = false; //조회화면 품목소분류 
                ddl_itemgp_select_s_amt.Visible = false;
                txt_to_currency1.Visible = false;
                txt_exchange.Visible = false;
                //list_select.Visible = false;


                sql_cust_cd = "SELECT SYSCODE_NAME cust_cd FROM SYSCODEDATA A WHERE  A.PLANT = 'CCUBEDIGITAL' AND A.SYSTABLE_NAME IN ( 'CUSTOMER') UNION ALL SELECT '%' FROM DUAL  ORDER BY 1 ";
                string sql = "";
                ds_sm_s4001_qty dt1 = new ds_sm_s4001_qty();

                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;//Request.QueryString["userid"];
                //  ReportViewer1.Reset();
                ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty.rdlc", "DataSet1");
                ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_amt.rdlc", "DataSet2");
                FillDropDownList(sql_cust_cd);
                //FillRadioButton();

                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_erp.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void FillDropDownList(string sql)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            try
            {
                // 품목 드랍다운리스트 내용을 보여준다.
                OracleCommand cmd2 = new OracleCommand(sql, conn);

                dr = cmd2.ExecuteReader();

            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void FillRadioButton()
        {
            string sql;
            //사용자 권한중 단가 권한이 있는지를 확인후 있으면 기준정보 입력화면을 셋팅해줌 없으면 집계 A /상세 B 만 ( 기준정보 C)
            sql = "select A.USR_ID  from z_usr_mast_rec_usr_role_asso  a inner join z_usr_mast_rec B ON A.USR_ID = B.USR_ID " +
                  " where usr_role_id like '%SA-PRICE00%' and USE_YN = 'y' and A.USR_ID = '" + Session["User"].ToString() + "' ";

            DataTable dt = Execute_ERP(sql);

            int chk_userid;
            chk_userid = dt.Rows.Count;
            if (chk_userid == 0)
            {
                rbl_view_type.Items.Add((new ListItem("기준정보등록", "A")));
                rbl_view_type.Items.Add((new ListItem("FCST 등록", "B")));
                rbl_view_type.Items.Add((new ListItem("FCST 조회", "C")));
            }
            else
            {
                rbl_view_type.Items.Add((new ListItem("기준정보등록", "A")));
                rbl_view_type.Items.Add((new ListItem("FCST 등록", "B")));
                rbl_view_type.Items.Add((new ListItem("FCST 조회", "C")));
            }
            rbl_view_type.SelectedIndex = 0;

        }
        private int Execute_ERP(SqlConnection connection, string sql, string wk_type)
        {
            connection.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;
            try
            {
                if (wk_type == "check")
                    value = Convert.ToInt32(cmd_erp.ExecuteScalar());
                else
                    value = cmd_erp.ExecuteNonQuery();
            }

            catch (Exception e)
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                value = -1;
            }

            connection.Close();
            return value;
        }

        private DataTable Execute_ERP(string sql)
        {
            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;
            DataTable dt = new DataTable();
            try
            {
                // 품목 드랍다운리스트 내용을 보여준다.
                //SqlConnection cmd2 = new SqlConnection(conn_erp);

                dr_erp = cmd_erp.ExecuteReader();
                dt.Load(dr_erp);
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();

            }
            conn_erp.Close();
            return dt;
        }
        private void ReportCreator(DataSet _dataSet, string sql, ReportViewer ReportViewer1, string _ReportName, string _ReportDataSourceName)
        {

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;

            DataSet ds = _dataSet;
            try
            {
                cmd_erp.CommandText = sql;
                dr_erp = cmd_erp.ExecuteReader();
                ds.Tables[0].Load(dr_erp);
                dr_erp.Close();
                ReportViewer1.LocalReport.ReportPath = Server.MapPath(_ReportName);

                ReportViewer1.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                ReportViewer1.LocalReport.DataSources.Add(rds);
                ReportViewer1.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }

        }

        private void ReportCreator2(DataSet _dataSet, string sql, ReportViewer ReportViewer2, string _ReportName, string _ReportDataSourceName)
        {

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;

            DataSet ds = _dataSet;
            try
            {
                cmd_erp.CommandText = sql;
                dr_erp = cmd_erp.ExecuteReader();
                ds.Tables[0].Load(dr_erp);
                dr_erp.Close();
                ReportViewer2.LocalReport.ReportPath = Server.MapPath(_ReportName);

                ReportViewer1.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                ReportViewer2.LocalReport.DataSources.Add(rds);
                ReportViewer2.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }

        }
        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (rbl_view_type.SelectedValue == "A") //기준정보등록 선택
            {
                //Panel_bas_info.Visible = true; //기준등록정보panel
                //rbl_bas_type.Visible = true; // 기준정보등록 라디오버튼
                Panel_Spread_Btn.Visible = true; //스프레드시트 Panel
                Panel_qty_amt.Visible = false;
                rdl_qty_amt.Visible = false;
                panel_upload.Visible = false;
                Panel_Spread_bas.Visible = true;
                //Panel_routeset.Visible = false;
                Panel_select.Visible = false;



            }
            if (rbl_view_type.SelectedValue == "B") //FCST 등록 선택
            {
                //Panel_bas_info.Visible = false;
                //rbl_bas_type.Visible = false;
                Panel_Spread_Btn.Visible = false;
                Panel_qty_amt.Visible = true;
                rdl_qty_amt.Visible = true;
                panel_upload.Visible = true;
                Panel_Spread_bas.Visible = false;
                //Panel_routeset.Visible = false;
                Panel_select.Visible = false;
                Panel_regist_excel_grid.Visible = true;
                // Label18.Visible = true;
                //Label5.Visible = false;
                //if (rbl_view_type.Items.tex)
                if (rdl_qty_amt.Items.FindByText("예상매출") != null)
                {
                    rdl_qty_amt.Items.FindByText("예상매출").Text = "단가";
                }

            }

            if (rbl_view_type.SelectedValue == "C")//FCST 조회 선택
            {
                //Panel_bas_info.Visible = false;
                //rbl_bas_type.Visible = false;
                Panel_Spread_Btn.Visible = false;
                Panel_qty_amt.Visible = true;
                rdl_qty_amt.Visible = true;
                panel_upload.Visible = false;
                Panel_Spread_bas.Visible = false;
                //Panel_routeset.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Label19.Visible = false;
                ddl_itemgp_select_amt.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;
                Panel_select_excel_amt_grid.Visible = true;
                if (rdl_qty_amt.Items.FindByText("단가") != null)
                {
                    rdl_qty_amt.Items.FindByText("단가").Text = "예상매출";
                }
                // Label18.Visible = true;
                //Label5.Visible = false;

            }
            ReportViewer1.Reset();
            ReportViewer2.Reset();
        }

        protected void btn_Add_Click(object sender, EventArgs e)
        {
            //if (rbl_bas_type.SelectedValue == "A")
            //{
            if (tb_rowcnt.Text == null || tb_rowcnt.Text == "")
            {
                MessageBox.ShowMessage("추가할Row수를입력해주세요.", this.Page);
                tb_rowcnt.Focus();
                return;
            }
            else
            {
                FpSpread1_ITEMGR.Sheets[0].AddRows(FpSpread1_ITEMGR.Sheets[0].RowCount, Convert.ToInt16(tb_rowcnt.Text));
            }
            //}
        }

        protected void btn_save_Click(object sender, EventArgs e)
        {
            // 품목그룹등록을 선택했으면
            //if (rbl_bas_type.SelectedValue == "A")
            //{
            FpSpread1_ITEMGR.SaveChanges();
            MessageBox.ShowMessage("저장되었습니다.", this.Page);
            //}
        }

        //MODIFY BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
        protected void FpSpread1_ITEMGR_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            //int colcnt;
            int i = 0;
            int r = (int)e.CommandArgument;
            int colcnt = e.EditValues.Count;

            string sql;

            //업데이트시
            if (FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value != null)
            {

                /*기존값가져오기*/
                string L_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value.ToString();
                string M_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 1].Value.ToString();
                string S_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 2].Value.ToString();
                string REMARK = FpSpread1_ITEMGR.Sheets[0].Cells[r, 3].Value.ToString();
                string cg_L_ITEM_AMT_GROUP, cg_M_ITEM_AMT_GROUP, cg_S_ITEM_AMT_GROUP, cg_REMARK;

                /*변경된값가져오기*/
                if (e.EditValues[0].ToString() == "System.Object")
                {
                    cg_L_ITEM_AMT_GROUP = L_ITEM_AMT_GROUP;
                }
                else
                {
                    cg_L_ITEM_AMT_GROUP = e.EditValues[0].ToString();
                }

                if (e.EditValues[1].ToString() == "System.Object")
                {
                    cg_M_ITEM_AMT_GROUP = M_ITEM_AMT_GROUP;
                }
                else
                {
                    cg_M_ITEM_AMT_GROUP = e.EditValues[1].ToString();
                }

                if (e.EditValues[2].ToString() == "System.Object")
                {
                    cg_S_ITEM_AMT_GROUP = S_ITEM_AMT_GROUP;
                }
                else
                {
                    cg_S_ITEM_AMT_GROUP = e.EditValues[2].ToString();
                }

                if (e.EditValues[3].ToString() == "System.Object" || e.EditValues[3].ToString() == "")
                {
                    cg_REMARK = " ";
                }
                else
                {
                    cg_REMARK = e.EditValues[3].ToString();
                }

                sql = "UPDATE T_DEVICE_AMT_GROUP_ADD SET L_ITEM_AMT_GROUP = '" + cg_L_ITEM_AMT_GROUP + "',"
                                + "M_ITEM_AMT_GROUP = '" + cg_M_ITEM_AMT_GROUP + "',"
                                + "S_ITEM_AMT_GROUP = '" + cg_S_ITEM_AMT_GROUP + "',"
                                + "REMARK = '" + cg_REMARK + "',"
                                + "UPDT_USER_ID = 'yoosr',"
                                + "UPDT_DT = SYSDATE "
                                + "WHERE L_ITEM_AMT_GROUP = '" + L_ITEM_AMT_GROUP + "'"
                                + "AND M_ITEM_AMT_GROUP = '" + M_ITEM_AMT_GROUP + "'"
                                + "AND S_ITEM_AMT_GROUP = '" + S_ITEM_AMT_GROUP + "'";

                QueryExecute(conn_if, sql, "");

                //r = r + 1;
            }
            else
            {
                string L_ITEM_AMT_GROUP = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
                string M_ITEM_AMT_GROUP = e.EditValues[1].ToString();
                string S_ITEM_AMT_GROUP = e.EditValues[2].ToString();
                string REMARK = e.EditValues[3].ToString();

                if (REMARK == "System.Object" || REMARK == "")
                    REMARK = " ";
                if (L_ITEM_AMT_GROUP == null || L_ITEM_AMT_GROUP == "")
                    MessageBox.ShowMessage("품목그룹명입력해주세요.", this.Page);
                else
                {
                    sql = "INSERT INTO T_DEVICE_AMT_GROUP_ADD "
                          + "VALUES('" + L_ITEM_AMT_GROUP + "','" + M_ITEM_AMT_GROUP + "',"
                                       + "'" + S_ITEM_AMT_GROUP + "','" + REMARK + "','yoosr',sysdate,'yoosr',sysdate)";

                    QueryExecute(conn_if, sql, "");
                    //r = r + 1;
                }
            }
        }

        //for (i = 0; i <= colcnt; i++)
        //{
        //    if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
        //    {
        //        string sql;

        //        //업데이트시
        //        if (FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value != null)
        //        {
        //            /*기존값가져오기*/
        //            string ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value.ToString();
        //            string REMARK = FpSpread1_ITEMGR.Sheets[0].Cells[r, 1].Value.ToString();
        //            string cg_ITEM_AMT_GROUP, cg_REMARK;

        //            /*변경된값가져오기*/
        //            if (i == 0)
        //                cg_ITEM_AMT_GROUP = e.EditValues[0].ToString();
        //            else
        //                cg_ITEM_AMT_GROUP = ITEM_AMT_GROUP;
        //            if (i == 1)
        //                cg_REMARK = e.EditValues[1].ToString();
        //            else
        //                cg_REMARK = REMARK;

        //            if (cg_REMARK == "System.Object" || cg_REMARK == "")
        //                cg_REMARK = " ";
        //            sql = "update T_DEVICE_AMT_GROUP ";
        //            sql = sql + "set ITEM_AMT_GROUP = '" + cg_ITEM_AMT_GROUP + "',REMARK = '" + cg_REMARK + "',updt_dt = sysdate ";
        //            sql = sql + " where ITEM_AMT_GROUP = '" + ITEM_AMT_GROUP + "'  ";
        //            QueryExecute(conn_if, sql, "");
        //        }
        //        else
        //        {
        //            //r = r + 1;
        //            //int j = FpSpread1.Sheets[0].ColumnCount;
        //            //string ITEM_AMT_GROUP = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
        //            //string REMARK = e.EditValues[1].ToString();

        //            string L_ITEM_AMT_GROUP = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
        //            string M_ITEM_AMT_GROUP = e.EditValues[1].ToString();
        //            string S_ITEM_AMT_GROUP = e.EditValues[2].ToString();
        //            string REMARK = e.EditValues[3].ToString();

        //            if (REMARK == "System.Object" || REMARK == "")
        //                REMARK = " ";
        //            //if (ITEM_AMT_GROUP == null || ITEM_AMT_GROUP == "")
        //            //    MessageBox.ShowMessage("품목그룹명입력해주세요.", this.Page);
        //            if (L_ITEM_AMT_GROUP == null || L_ITEM_AMT_GROUP == "")
        //                MessageBox.ShowMessage("품목그룹명입력해주세요.", this.Page);
        //            else
        //            {
        //                //sql = "insert into T_DEVICE_AMT_GROUP ";
        //                //sql = sql + "values('" + ITEM_AMT_GROUP + "','" + REMARK + "', 'unierp', sysdate, 'unierp', sysdate)";

        //                sql = "INSERT INTO T_DEVICE_AMT_GROUP_ADD "
        //                      + "VALUES('" + L_ITEM_AMT_GROUP + "','" + M_ITEM_AMT_GROUP + "',"
        //                                   + "'" + S_ITEM_AMT_GROUP + "','" + REMARK + "','yoosr',sysdate,'yoosr',sysdate)";

        //                QueryExecute(conn_if, sql, "");

        //                r = r + 1;

        //            }
        //        }
        //    }
        //}

        public int QueryExecute(OracleConnection connection, string sql, string wk_type)
        {

            conn_if.Open();
            cmd = conn_if.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {

                if (wk_type == "check")
                    value = Convert.ToInt32(cmd.ExecuteScalar());
                else
                    value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (conn_if.State == ConnectionState.Open)
                    conn_if.Close();
                value = -1;
            }

            conn_if.Close();
            return value;
        }


        public DataTable QueryExeuteDT(string sql)
        {
            ds_sm_s4001 ds = new ds_sm_s4001();
            ds_sm_s4001_test ds1 = new ds_sm_s4001_test();

            conn_if.Open();
            cmd = conn_if.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                conn_if.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn_if.Close();
            }

            return ds.Tables[0];
        }

        //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
        //조회
        public void serch_list()
        {
            string sql;
            // 품목그룹 조회시
            //if (rbl_bas_type.SelectedValue == "A")
            //{
            //sql = "select ITEM_AMT_GROUP, REMARK from T_DEVICE_AMT_GROUP ";

            sql = "SELECT L_ITEM_AMT_GROUP,M_ITEM_AMT_GROUP,S_ITEM_AMT_GROUP,REMARK FROM T_DEVICE_AMT_GROUP_ADD ORDER BY 1,2";

            sqlAdapter1 = new OracleDataAdapter(sql, conn_if);

            //sqlAdapter1.Fill(ds, "ds");
            sqlAdapter1.Fill(ds1, "ds");


            FpSpread1_ITEMGR.DataSource = ds1;
            FpSpread1_ITEMGR.DataBind();
            //}

        }


        protected void btn_exe_Click(object sender, EventArgs e)
        {
            serch_list();
        }

        //MODIFY BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
        protected void btn_Delete_Click(object sender, EventArgs e)
        {

            System.Collections.IEnumerator enu = FpSpread1_ITEMGR.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;

            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread1_ITEMGR.Sheets[0].ActiveRow;

                string L_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 0].Text;
                string M_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1].Text;
                string S_ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 2].Text;

                string sql = "DELETE FROM T_DEVICE_AMT_GROUP_ADD "
                             + "WHERE L_ITEM_AMT_GROUP = '" + L_ITEM_AMT_GROUP + "'"
                             + " AND M_ITEM_AMT_GROUP = '" + M_ITEM_AMT_GROUP + "'"
                             + " AND S_ITEM_AMT_GROUP = '" + S_ITEM_AMT_GROUP + "'";

                if (QueryExecute(conn_if, sql, "") > 0)
                    FpSpread1_ITEMGR.Sheets[0].Rows.Remove(FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1);
            }

            MessageBox.ShowMessage("삭제되었습니다.", this.Page);



            //if (rbl_bas_type.SelectedValue == "A")
            //{
            //    System.Collections.IEnumerator enu = FpSpread1_ITEMGR.ActiveSheetView.SelectionModel.GetEnumerator();
            //    FarPoint.Web.Spread.Model.CellRange cr;

            //    while (enu.MoveNext())
            //    {
            //        cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
            //        int a = FpSpread1_ITEMGR.Sheets[0].ActiveRow;
            //        //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
            //        for (int i = 0; i < cr.RowCount; i++)
            //        {
            //            string ITEM_AMT_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 0].Text;
            //            string REMARK = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1].Text;


            //            string sql = "delete T_DEVICE_AMT_GROUP ";
            //            sql = sql + " where ITEM_AMT_GROUP  ='" + ITEM_AMT_GROUP + "' ";

            //            if (QueryExecute(conn_if, sql, "") > 0)
            //                FpSpread1_ITEMGR.Sheets[0].Rows.Remove(FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1);
            //        }
            //    }
            //}

            //MessageBox.ShowMessage("삭제되었습니다.", this.Page);
        }
        //protected void rbl_bas_type_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    //1. 소분류품목그룹등록
        //    if (rbl_bas_type.SelectedValue == "A")
        //    {
        //        Panel_Spread_bas.Visible = true;
        //        Panel_routeset.Visible = false;
        //        panel_upload.Visible = false;
        //        Panel_qty_amt.Visible = false;
        //        Panel_select.Visible = false;

        //    }
        //    //2. 대분류-소분류 품목그룹연결
        //    if (rbl_bas_type.SelectedValue == "B")
        //    {
        //        Panel_Spread_bas.Visible = false;
        //        Panel_routeset.Visible = true;
        //        panel_upload.Visible = false;
        //        Panel_qty_amt.Visible = false;
        //        Panel_select.Visible = false;


        //    }
        //    ReportViewer1.Reset();
        //}

        protected void rbl_qty_amt_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            //
            if (rdl_qty_amt.SelectedValue == "A") //수량선택
            {
                Label16.Visible = true; //등록화면 월 선택 (보임)
                txt_regist_date_mm.Visible = true;

                Label7.Visible = true;//조회화면 월선택 (보임)
                txt_select_date_mm.Visible = true;

                Label19.Visible = false; //조회화면 품목중분류 
                ddl_itemgp_select_amt.Visible = false;

                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
                Label20.Visible = false; //조회화면 품목소분류 
                ddl_itemgp_select_s_amt.Visible = false;

                txt_to_currency1.Visible = false;
                txt_exchange.Visible = false;

                //list_select.Visible = false;

                Panel_select_excel_qty_grid.Visible = true;
                ReportViewer1.Visible = true;

            }

            else if (rdl_qty_amt.SelectedValue == "B")//금액선택
            {
                Label16.Visible = true;//등록화면 월 선택 (보임)
                txt_regist_date_mm.Visible = true;

                Label7.Visible = false;//조회화면 월 선택 (보이지 않음)
                txt_select_date_mm.Visible = false;

                Label19.Visible = true;//조회화면 품목소분류  (보임)
                ddl_itemgp_select_amt.Visible = true;

                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
                Label20.Visible = true; //조회화면 품목소분류 
                ddl_itemgp_select_s_amt.Visible = true;

                txt_to_currency1.Visible = true;
                txt_exchange.Visible = true;

                //list_select.Visible = true;

                Panel_select_excel_amt_grid.Visible = true;
                ReportViewer2.Visible = true;

            }
            ReportViewer1.Reset();
            ReportViewer2.Reset();
        }






        //DELETE BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
        //***************************대분류-소분류 품목그룹 연결화면- 조회버튼******************************
        //protected void btn_exe_itemgp_routeset_Click(object sender, EventArgs e)
        //{
        //    string sql;
        //    lsb_l_routeset.Items.Clear(); //내용지우기
        //    lsb_r_routeset.Items.Clear(); //내용지우기

        //    if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
        //    {
        //        MessageBox.ShowMessage("대분류 품목그룹을선택해주세요.", this.Page);
        //    }
        //    else
        //    {
        //        //왼쪽 미등록소분류품목 가져오기
        //        sql = "select DISTINCT ITEM_AMT_GROUP " +
        //              "from T_device_amt_group where item_amt_group NOT IN (SELECT Item_amt_group FROM T_device_group_amt_group where ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "' ) ";

        //        DataTable dt = QueryExeuteDT(sql);
        //        if (dt.Rows.Count > 0)
        //        {
        //            lsb_l_routeset.DataSource = dt;
        //            lsb_l_routeset.DataTextField = "ITEM_AMT_GROUP";
        //            lsb_l_routeset.DataValueField = "ITEM_AMT_GROUP";
        //            lsb_l_routeset.DataBind();
        //        }
        //        //오른쪽 등록되어있는 소분류품복 가져오기
        //        sql = "select Item_amt_group from T_device_group_amt_group WHERE ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "'  ";
        //        dt = QueryExeuteDT(sql);
        //        if (dt.Rows.Count > 0)
        //        {
        //            lsb_r_routeset.DataSource = dt;
        //            lsb_r_routeset.DataTextField = "Item_amt_group";
        //            lsb_r_routeset.DataValueField = "Item_amt_group";
        //            lsb_r_routeset.DataBind();
        //        }
        //    }
        //}

        //protected void btn_move_right_Click(object sender, EventArgs e)
        //{

        //    if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
        //    {
        //        MessageBox.ShowMessage("품목그룹을선택해주세요.", this.Page);
        //    }
        //    else
        //    {
        //        for (int i = 0; i < lsb_l_routeset.Items.Count; i++)
        //        {
        //            if (this.lsb_l_routeset.Items[i].Selected)
        //            {
        //                this.lsb_r_routeset.Items.Add(this.lsb_l_routeset.Items[i]);

        //                //선택된품목그룹에왼쪽리스트박스내용을insert한다. 
        //                string sql = "insert into T_device_group_amt_group " +
        //                             "values ('" + ddl_itemgp.SelectedValue + "', '" + this.lsb_l_routeset.Items[i] + "','" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate ) ";

        //                if (QueryExecute(conn_if, sql, "") <= 0)
        //                    MessageBox.ShowMessage("데이타저장에실패했습니다.", this.Page);
        //                this.lsb_l_routeset.Items.Remove(this.lsb_l_routeset.Items[i]);
        //                i--;
        //            }
        //        }

        //        btn_exe_itemgp_routeset_Click(null, null);
        //    }
        //}

        //protected void btn_move_left_Click(object sender, EventArgs e)
        //{
        //    if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
        //    {
        //        MessageBox.ShowMessage("품목그룹을선택해주세요.", this.Page);
        //    }
        //    else
        //    {
        //        for (int i = 0; i < lsb_r_routeset.Items.Count; i++)
        //        {
        //            if (this.lsb_r_routeset.Items[i].Selected)
        //            {
        //                this.lsb_l_routeset.Items.Add(this.lsb_r_routeset.Items[i]);
        //                //선택된품목그룹에왼쪽리스트박스내용을insert한다. 
        //                string sql = "delete T_device_group_amt_group " +
        //                             "where ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "' and item_amt_group =  '" + this.lsb_r_routeset.Items[i] + "'";

        //                if (QueryExecute(conn_if, sql, "") <= 0)
        //                    MessageBox.ShowMessage("데이타저장에실패했습니다.", this.Page);

        //                this.lsb_r_routeset.Items.Remove(this.lsb_r_routeset.Items[i]);
        //                i--;

        //            }
        //        }
        //        btn_exe_itemgp_routeset_Click(null, null);
        //    }

        //}
        /*********************************엑셀업로드 ************************************/

        // 엑셀업로드 클릭
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (list_regist_version.SelectedValue.ToString() == "-선택안함-" || list_regist_version.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                return;
            }

            if (txt_regist_date_yyyy.Text == "" || txt_regist_date_yyyy.Text == null)
            {
                MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
                return;
            }

            if (rdl_qty_amt.SelectedValue.ToString() == "A" && //수량선택일 경우는 '월'까지 선택
                (txt_regist_date_mm.Text == "" || txt_regist_date_mm.Text == null))
            {
                MessageBox.ShowMessage("월을선택해주세요.", this.Page);
                return;
            }
            if (FileUpload1.HasFile)
            {
                string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName).ToUpper();
                string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                string FilePath = Server.MapPath(FolderPath + FileName);
                FileUpload1.SaveAs(FilePath);
                if (Extension == ".XLS" || Extension == ".XLSX")
                    GetExcelSheets(FilePath, Extension, "Yes");
                else
                    MessageBox.ShowMessage("Excel 파일만 업로드 가능합니다", this);
            }
        }




        // 엑셀sheet 받아오기
        private void GetExcelSheets(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".XLS": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".XLSX": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                             .ConnectionString;
                    break;
            }

            //Get the Sheets in Excel WorkBoo
            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;
            connExcel.Open();

            //Bind the Sheets to DropDownList
            ddlSheets.Items.Clear();
            ddlSheets.Items.Add(new ListItem("--Select Sheet--", ""));
            ddlSheets.DataSource = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            ddlSheets.DataTextField = "TABLE_NAME";
            ddlSheets.DataValueField = "TABLE_NAME";

            ddlSheets.DataBind();

            DataTable dtCSV = new DataTable();
            OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + ddlSheets.SelectedValue + "]", connExcel);
            OleDbDataAdapter daCSV = new OleDbDataAdapter();
            // OleDbCommand DBCommand;
            IDataReader DBReader;
            DBReader = cmdSelect.ExecuteReader();
            grid_regist_excel.DataSource = DBReader;
            grid_regist_excel.DataBind();

            DBReader.Close();
            connExcel.Close();
            HiddenField_fileName.Value = Path.GetFileName(FilePath); //파일명 저장용
            HiddenField_filePath.Value = FilePath; //파일경로 저장용
            HiddenField_extension.Value = Extension; //파일확장자 저장용



            grid_regist_excel.Visible = true; //그리드뷰 보여주기
        }

        // 엑셀sheet 선택시 보여주기 위한 함수
        private void ViewExcelSheets(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".XLS": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".XLSX": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                             .ConnectionString;
                    break;
            }

            //Get the Sheets in Excel WorkBoo
            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;
            connExcel.Open();

            DataTable dtCSV = new DataTable();
            OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + ddlSheets.SelectedValue + "]", connExcel);
            OleDbDataAdapter daCSV = new OleDbDataAdapter();
            // OleDbCommand DBCommand;
            IDataReader DBReader;
            DBReader = cmdSelect.ExecuteReader();
            grid_regist_excel.DataSource = DBReader;
            grid_regist_excel.DataBind();

            DBReader.Close();
            connExcel.Close();

        }


        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            if (rdl_qty_amt.SelectedValue.ToString() == "A") //수량화면 선택시
            {

                InsertFCST_QTY();
            }
            else
            {
                InsertFCST_AMT();
            }
        }
        /*********************************수량 테이블 저장 ************************************/


        private void InsertFCST_QTY()
        {
            bool bSuccess = false;

            for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
            {
                chk_save_yn = 0;
                // FpSpread_select_amt.Sheets[0].Cells[i, 0].Value = list_select_version.SelectedValue.ToString();
                // FpSpread_select_amt.Sheets[0].Cells[i, j + 1].Value = grid_regist_excel.Rows[j].Cells[i].Text.Trim();


                string version_no = list_regist_version.SelectedValue.ToString();
                string cust_nm = grid_regist_excel.Rows[i].Cells[0].Text.Trim();//거래처
                string item_nm = grid_regist_excel.Rows[i].Cells[1].Text.Trim();//품목명
                string item_gp = grid_regist_excel.Rows[i].Cells[2].Text.Trim();//품목대분류
                string size = grid_regist_excel.Rows[i].Cells[3].Text.Trim();//wafer size
                string process_type = grid_regist_excel.Rows[i].Cells[4].Text.Trim(); //process type
                string route = grid_regist_excel.Rows[i].Cells[5].Text.Trim();//route
                string pkg_type = grid_regist_excel.Rows[i].Cells[6].Text.Trim(); //pkg type
                string bas_yyyymm = txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString(); //적용년월(화면에서 년도선택, 월선택)
                //string remark = grid_regist_excel.Rows[i].Cells[13].Text.Trim(); //remark
                // string plan_mm = txt_regist_date_mm.SelectedValue.ToString();//계획월 수량시작월은 화면에서 선택한 월이고 그다음 칼럼부터는 + 1 월을 해준다 (ex)Plan_nm='02')

                string selectDate = txt_regist_date_mm.SelectedValue.ToString() + "/01/" + txt_regist_date_yyyy.Text;
                DateTime dtTmp = Convert.ToDateTime(selectDate);
                for (int j = 0; j < 6; j++)
                {
                    if (j > 0)
                        dtTmp = dtTmp.AddMonths(1);

                    string sMonth = dtTmp.Month.ToString("00");
                    string sYear = dtTmp.Year.ToString();

                    string qty1 = grid_regist_excel.Rows[i].Cells[j + 7].Text.Trim();

                    string sql = "select count(item_nm) from S_FCST_QTY_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'and cust_nm = '" + cust_nm + "' ";
                    sql += "and item_nm = '" + item_nm + "' ";
                    sql += "and item_gp = '" + item_gp + "' and bas_yyyymm = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' and plan_mm= '" + sYear + sMonth + "' ";
                    //sql += "and qty = '" + qty1 + "' and remark = '" + remark + "'";

                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {

                        sql = "delete S_FCST_QTY_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'and cust_nm = '" + cust_nm + "' ";
                        sql += "and item_nm = '" + item_nm + "' ";
                        sql += "and item_gp = '" + item_gp + "' and bas_yyyymm = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' and plan_mm= '" + sYear + sMonth + "' ";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                        {
                            sql = "insert into S_FCST_QTY_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + item_gp + "' ,'" + size + "' ,'" + process_type + "',";
                            sql += "'" + route + "' ,'" + pkg_type + "' , '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' ,'" + sYear + sMonth + "','" + qty1 + "',";
                            sql += "'','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((list_regist_version.SelectedValue.ToString() == "" || list_regist_version.SelectedValue.ToString() == null || list_regist_version.SelectedValue.ToString() == "&nbsp;") || (cust_nm == "" || cust_nm == null || cust_nm == "&nbsp;")
                            || (item_nm == "" || item_nm == null || item_nm == "&nbsp;") || (item_gp == "" || item_gp == null || item_gp == "&nbsp;")
                            || (size == "" || size == null || size == "&nbsp;") || (process_type == "" || process_type == null || process_type == "&nbsp;")//값체크
                            || (route == "" || route == null || route == "&nbsp;") || (pkg_type == "" || pkg_type == null || item_gp == "&nbsp;")
                            || (bas_yyyymm == "" || bas_yyyymm == null || bas_yyyymm == "&nbsp;") || (sMonth == "" || sMonth == null || sMonth == "&nbsp;")
                            || (qty1 == "" || qty1 == null || qty1 == "&nbsp;"))// || (remark == "" || remark == null || remark == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;

                        }

                        else
                        {

                            //if (grid_regist_excel.Rows.Count == i)
                            //{
                            sql = "insert into S_FCST_QTY_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + item_gp + "' ,'" + size + "' ,'" + process_type + "',";
                            sql += "'" + route + "' ,'" + pkg_type + "' , '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' ,'" + sYear + sMonth + "','" + qty1 + "',";
                            sql += "'','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";
                            //}
                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }


                }

                if (chk_save_yn == 6)
                {
                    bSuccess = true;
                }
                else
                {
                    bSuccess = false;
                    break;
                }
            }

            if (bSuccess == true)
            {
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            else
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);
            }


            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }

        /***********************금액 업로드파일 저장************************************************/
        private void InsertFCST_AMT()
        {
            bool bSuccess = false;

            for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
            {
                chk_save_yn = 0;


                string version_no = list_regist_version.SelectedValue.ToString();
                string cust_nm = grid_regist_excel.Rows[i].Cells[0].Text.Trim();//거래처
                string item_nm = grid_regist_excel.Rows[i].Cells[1].Text.Trim();//품목명
                string item_gp = grid_regist_excel.Rows[i].Cells[2].Text.Trim();//품목대분류
                string item_amt_gp = grid_regist_excel.Rows[i].Cells[3].Text.Trim(); //품목소분류 선택
                string size = grid_regist_excel.Rows[i].Cells[4].Text.Trim();//wafer size
                string process_type = grid_regist_excel.Rows[i].Cells[5].Text.Trim(); //process type
                string route = grid_regist_excel.Rows[i].Cells[6].Text.Trim();//route
                string pkg_type = grid_regist_excel.Rows[i].Cells[7].Text.Trim(); //pkg type
                string plan_curr_unit = grid_regist_excel.Rows[i].Cells[8].Text.Trim();
                string bas_yyyy = txt_regist_date_yyyy.Text; //적용년도(화면에서 년도만선택)
                //string remark = grid_regist_excel.Rows[i].Cells[15].Text.Trim(); //remark
                string selectDate = txt_regist_date_mm.SelectedValue.ToString() + "/01/" + txt_regist_date_yyyy.Text;
                DateTime dtTmp = Convert.ToDateTime(selectDate);
                for (int j = 0; j < 6; j++)
                {
                    if (j > 0)
                        dtTmp = dtTmp.AddMonths(1);

                    string sMonth = dtTmp.Month.ToString("00");
                    string sYear = dtTmp.Year.ToString();

                    string plan_amt = grid_regist_excel.Rows[i].Cells[j + 9].Text.Trim();


                    string sql = "select count(item_nm) from S_FCST_AMT_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'and cust_nm = '" + cust_nm + "' ";
                    sql += "and item_nm = '" + item_nm + "' ";
                    sql += "and item_gp = '" + item_gp + "' and item_amt_gp= '" + item_amt_gp + "'and bas_yyyy = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' and plan_mm= '" + sYear + sMonth + "' ";


                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {

                        sql = "delete S_FCST_AMT_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'and cust_nm = '" + cust_nm + "' ";
                        sql += "and item_nm = '" + item_nm + "' ";
                        sql += "and item_gp = '" + item_gp + "'and item_amt_gp= '" + item_amt_gp + "'and bas_yyyy = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "' and plan_mm= '" + sYear + sMonth + "' ";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                        {
                            sql = "insert into S_FCST_AMT_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + item_gp + "' ,'" + size + "' ,'" + process_type + "',";
                            sql += "'" + route + "' ,'" + pkg_type + "' , '" + item_amt_gp + "','" + txt_regist_date_yyyy.Text + "' ,'" + sYear + sMonth + "','" + plan_amt + "','" + plan_curr_unit + "',";
                            sql += "'','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";
                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((list_regist_version.SelectedValue.ToString() == "" || list_regist_version.SelectedValue.ToString() == null || list_regist_version.SelectedValue.ToString() == "&nbsp;") || (cust_nm == "" || cust_nm == null || cust_nm == "&nbsp;")
                            || (item_nm == "" || item_nm == null || item_nm == "&nbsp;") || (item_gp == "" || item_gp == null || item_gp == "&nbsp;")
                            || (size == "" || size == null || size == "&nbsp;") || (process_type == "" || process_type == null || process_type == "&nbsp;")
                            || (route == "" || route == null || route == "&nbsp;") || (pkg_type == "" || pkg_type == null || item_gp == "&nbsp;")
                            || (bas_yyyy == "" || bas_yyyy == null || bas_yyyy == "&nbsp;") || (plan_curr_unit == "" || plan_curr_unit == null || plan_curr_unit == "&nbsp;")
                            || (item_amt_gp == "" || item_amt_gp == null || item_amt_gp == "&nbsp;") || (sMonth == "" || sMonth == null || sMonth == "&nbsp;")
                            || (plan_amt == "" || plan_amt == null || plan_amt == "&nbsp;"))//|| (remark == "" || remark == null || remark == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;
                        }
                        else
                        {

                            //if (grid_regist_excel.Rows.Count == i)
                            //{
                            sql = "insert into S_FCST_AMT_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + item_gp + "' ,'" + size + "' ,'" + process_type + "',";
                            sql += "'" + route + "' ,'" + pkg_type + "' , '" + item_amt_gp + "','" + txt_regist_date_yyyy.Text + "' ,'" + sYear + sMonth + "','" + plan_amt + "','" + plan_curr_unit + "',";
                            sql += "'','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";
                            //}
                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }


                }

                if (chk_save_yn == 6)
                {
                    bSuccess = true;
                }
                else
                {
                    bSuccess = false;
                    break;
                }
            }

            if (bSuccess == true)
            {
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            else
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);
            }

            ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }


        protected void btnCancel_Click(object sender, EventArgs e)
        {
            //그리드 뷰를 초기화 한다.
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
            grid_regist_excel.Visible = false;
        }


        /******************************* 조회 *************************************************************************************/


        protected void btn_select_Click(object sender, EventArgs e) //조회버튼 클릭
        {
            if (rdl_qty_amt.SelectedValue.ToString() == "A")
            {
                ReportViewer1.Reset();

                if (list_select_version.SelectedValue.ToString() == "-선택안함-" || list_select_version.SelectedValue.ToString() == null)
                {
                    MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                    return;
                }
                if (txt_select_date_yyyy.Text == "" || txt_select_date_yyyy.Text == null)
                {
                    MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
                    return;
                }
                if (rdl_qty_amt.SelectedValue.ToString() == "A" && //수량선택일 경우는 '월'까지 선택
                    (txt_select_date_mm.Text == "" || txt_select_date_mm.Text == null))
                {
                    MessageBox.ShowMessage("월을선택해주세요.", this.Page);
                    return;
                }

                //if (ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null)
                //{
                //    MessageBox.ShowMessage("품목대분류를 선택해주세요.", this.Page);
                //    ReportViewer1.Reset();
                //    return;
                //}

                ReportViewer1.Reset();
                string sql = "SELECT distinct cust_nm,item_nm,item_gp,size,process_type,route,pkg_type,(SUBSTRING(plan_mm,1,4)+'.'+SUBSTRING(plan_mm,5,6))AS plan_mm,qty FROM S_FCST_QTY_IMPORT";//수량쿼리실행                            
                sql += " where bas_yyyymm = '" + txt_select_date_yyyy.Text + txt_select_date_mm.Text + "'";//bas_yyyymm 과 선택한 기준년월은 동일               
                sql += " and version_no =  '" + list_select_version.SelectedValue.ToString() + "'";//선택한 버전에 있는것만
                if (!(ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null))
                {
                    sql += " and item_gp= '" + ddl_itemgp_select.SelectedValue + "'"; //품목대그룹은 선택된 품목대그룹과 동일해야함
                }
                sql += " order by cust_nm,item_nm";//고객사, 디바이스 순 정렬

                ReportViewer1.Reset();
                ds_sm_s4001_qty dt1 = new ds_sm_s4001_qty();

                //ADD BY SLYOO : FCST 추가 개발 WITH 박세진 : 20140910
                if (list_select.Text == "고객사")
                {
                    ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty_cus.rdlc", "DataSet1");
                }
                else if (list_select.Text == "대분류")
                {
                    ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty_item.rdlc", "DataSet1");
                }
                else
                {
                    ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty.rdlc", "DataSet1");
                }


                //ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty.rdlc", "DataSet1");

            }
            if (rdl_qty_amt.SelectedValue.ToString() == "B")
            {
                Panel_select_excel_amt_grid.Visible = true;
                ReportViewer2.Reset();

                if (list_select_version.SelectedValue.ToString() == "-선택안함-" || list_select_version.SelectedValue.ToString() == null)
                {
                    MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                    return;
                }
                if (txt_select_date_yyyy.Text == "" || txt_select_date_yyyy.Text == null)
                {
                    MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
                    return;
                }

                //DELETE BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140722
                //if (rdl_qty_amt.SelectedValue.ToString() == "A" && //수량선택일 경우는 '월'까지 선택
                //    (txt_select_date_mm.Text == "" || txt_select_date_mm.Text == null))
                //{
                //    MessageBox.ShowMessage("월을선택해주세요.", this.Page);
                //    return;
                //}

                ////DELETE BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140722
                //if (ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null)
                //{
                //    MessageBox.ShowMessage("품목대분류를 선택해주세요.", this.Page);
                //    return;
                //}

                ////DELETE BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140722
                //if (ddl_itemgp_select_amt.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null)
                //{
                //    MessageBox.ShowMessage("품목소분류를 선택해주세요.", this.Page);
                //    return;
                //}
                ReportViewer2.Reset();

                string sql = "";

                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140722
                if (txt_to_currency1.Text == "" || txt_exchange.Text == "")
                {
                    sql += "SELECT distinct cust_nm,item_nm,item_gp,item_amt_gp,size,process_type,route,pkg_type,bas_yyyy,(SUBSTRING(plan_mm,1,4)+'.'+SUBSTRING(plan_mm,5,6))AS plan_mm,plan_amt,plan_curr_unit,remark";//금액쿼리실행
                }
                else
                {
                    sql += "SELECT DISTINCT CUST_NM,ITEM_NM,ITEM_GP,ITEM_AMT_GP,SIZE,PROCESS_TYPE,ROUTE,PKG_TYPE,BAS_YYYY,(SUBSTRING(plan_mm,1,4)+'.'+SUBSTRING(plan_mm,5,6))AS plan_mm,PLAN_AMT,PLAN_CURR_UNIT,CAST(ROUND(PLAN_AMT*'" + txt_exchange.Text + "',5,1) AS DECIMAL(10,5)) AS EXCHANGE_PLANT_AMT,'" + txt_to_currency1.Text + "' AS EXCHANGE_UNIT,REMARK ";//금액쿼리실행
                }

                sql += " FROM S_FCST_AMT_IMPORT"; //품목대그룹은 선택된 품목대그룹과 동일해야함
                sql += " WHERE version_no =  '" + list_select_version.SelectedValue.ToString() + "'";//선택한 버전에 있는것만
                sql += " AND bas_yyyy='" + txt_select_date_yyyy.Text + "'";

                if (!(ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == ""))
                {
                    sql += " AND item_gp= '" + ddl_itemgp_select.SelectedValue + "'";
                }
                if (!(ddl_itemgp_select_amt.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select_amt.SelectedValue.ToString() == ""))
                {
                    sql += " AND route= '" + ddl_itemgp_select_amt.SelectedValue + "'";
                }
                if (!(ddl_itemgp_select_s_amt.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select_s_amt.SelectedValue.ToString() == ""))
                {
                    sql += " AND item_amt_gp='" + ddl_itemgp_select_s_amt.SelectedValue + "'";
                }

                ReportViewer2.Reset();
                ds_sm_s4001_amt dt1 = new ds_sm_s4001_amt();

                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140722
                if (txt_to_currency1.Text == "" || txt_exchange.Text == "")
                {

                    if (list_select.Text == "고객사")
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_cus.rdlc", "DataSet2");
                    }
                    else if (list_select.Text == "대분류")
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_item.rdlc", "DataSet2");
                    }
                    else
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_amt.rdlc", "DataSet2");
                    }

                }
                else
                {

                    if (list_select.Text == "고객사")
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_cus_add.rdlc", "DataSet1");
                    }
                    else if (list_select.Text == "대분류")
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_item_add.rdlc", "DataSet1");
                    }
                    else
                    {
                        ReportCreator2(dt1, sql, ReportViewer2, "rp_sm_s4001_amt_add.rdlc", "DataSet1");
                    }
                }

            }

        }


        protected void ddl_itemgp_select_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer2.Reset();
            string sql;
            ddl_itemgp_select_amt.Items.Clear(); //내용지우기
            ddl_itemgp_select_s_amt.Items.Clear();

            if (ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null)
            {
                return;
            }
            else
            {
                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
                sql = "SELECT DISTINCT M_ITEM_AMT_GROUP FROM T_DEVICE_AMT_GROUP_ADD WHERE L_ITEM_AMT_GROUP = '" + ddl_itemgp_select.SelectedValue + "'"
                      + "union all select '-선택안됨-' from dual where rownum < 2 order by 1";
                DataTable dt = QueryExeuteDT(sql);
                if (dt.Rows.Count > 0)
                {
                    ddl_itemgp_select_amt.DataSource = dt;
                    ddl_itemgp_select_amt.DataTextField = "M_ITEM_AMT_GROUP";
                    ddl_itemgp_select_amt.DataValueField = "M_ITEM_AMT_GROUP";
                    ddl_itemgp_select_amt.DataBind();
                }
            }
        }

        protected void ddl_itemgp_select_amt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer2.Reset();
            string sql;
            ddl_itemgp_select_s_amt.Items.Clear(); //내용지우기

            if (ddl_itemgp_select.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_select.SelectedValue.ToString() == null || ddl_itemgp_select_amt.SelectedValue.ToString() == null)
            {
                return;
            }
            else
            {
                //ADD BY SLYOO : FCST 추가 개발 WITH 박두순 : 20140717
                sql = "SELECT DISTINCT S_ITEM_AMT_GROUP FROM T_DEVICE_AMT_GROUP_ADD WHERE L_ITEM_AMT_GROUP = '" + ddl_itemgp_select.SelectedValue + "' "
                    + "AND M_ITEM_AMT_GROUP = '" + ddl_itemgp_select_amt.SelectedValue + "'"
                    + "union all select '-선택안됨-' from dual where rownum < 2 order by 1";
                DataTable dt = QueryExeuteDT(sql);
                if (dt.Rows.Count > 0)
                {
                    ddl_itemgp_select_s_amt.DataSource = dt;
                    ddl_itemgp_select_s_amt.DataTextField = "S_ITEM_AMT_GROUP";
                    ddl_itemgp_select_s_amt.DataValueField = "S_ITEM_AMT_GROUP";
                    ddl_itemgp_select_s_amt.DataBind();
                }
            }
        }

        protected void ddlSheets_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            ViewExcelSheets(HiddenField_filePath.Value, HiddenField_extension.Value, "Yes");
        }

        protected void ddl_exchange_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportViewer2.Reset();

            conn_erp.Open();

            string exchange = "";
            string to_currency = "";
            string from_currency = "";

            string sql = "SELECT "
                             + "C.STD_RATE,"
                             + "C.TO_CURRENCY,"
                             + "C.FROM_CURRENCY"
                        + " FROM B_CURRENCY A, B_CURRENCY B, B_DAILY_EXCHANGE_RATE C"
                        + " WHERE C.FROM_CURRENCY = A.CURRENCY"
                             + " AND C.TO_CURRENCY = B.CURRENCY"
                             + " AND A.CURRENCY IN (SELECT  DISTINCT PLAN_CURR_UNIT"
                                                             + " FROM S_FCST_AMT_IMPORT"
                                                             + " WHERE ITEM_GP= '" + ddl_itemgp_select.SelectedValue + "'"
                                                             + " AND ROUTE= '" + ddl_itemgp_select_amt.SelectedValue + "' "
                                                             + " AND ITEM_AMT_GP='" + ddl_itemgp_select_s_amt.SelectedValue + "'"
                                                             + " AND VERSION_NO =  '" + list_select_version.SelectedValue.ToString() + "' "
                                                             + " AND BAS_YYYY='" + txt_select_date_yyyy.Text + "')"
                             + " AND C.APPRL_DT = '" + DateTime.Now.ToShortDateString() + "'"
                             + "ORDER BY A.CURRENCY ASC, B.CURRENCY ASC";

            SqlCommand sm = new SqlCommand(sql, conn_erp);
            SqlDataReader sr = sm.ExecuteReader();

            while (sr.Read())
            {
                exchange = sr.GetSqlDecimal(0).ToString();
                to_currency = sr.GetString(1);
                from_currency = sr.GetString(2);
            }

            txt_exchange.Text = exchange;
            txt_to_currency1.Text = to_currency;

            sr.Close();
            conn_erp.Close();
        }

    }
}

