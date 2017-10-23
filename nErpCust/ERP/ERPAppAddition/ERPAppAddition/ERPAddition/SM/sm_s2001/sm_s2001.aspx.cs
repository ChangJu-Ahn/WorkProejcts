using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
//using System.Data.OleDb;
//using System.Data.OracleClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;


namespace ERPAppAddition.ERPAddition.SM.sm_s2001
{
    public partial class sm_s2001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];

        string sql_cust_cd;

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleConnection conn_if = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_UNIERP"].ConnectionString);

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        SqlDataReader sql_dr;

        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;
        OracleDataAdapter sqlAdapter1;
        DataSet ds = new DataSet();
        string sql_spread, sql;
        int value;
        protected void Page_Load(object sender, EventArgs e)
        {
            //fillbp();
            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void fillbp()
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            try
            {
                string sql_cust_cd = "SELECT SYSCODE_NAME cust_cd FROM SYSCODEDATA A WHERE  A.PLANT = 'CCUBEDIGITAL' AND A.SYSTABLE_NAME IN ( 'CUSTOMER') UNION ALL SELECT '%' FROM DUAL ORDER BY 1   ";
                OracleCommand cmd2 = new OracleCommand(sql_cust_cd, conn);

                dr = cmd2.ExecuteReader();
                // ListItem liObject = ddl_cust_cd.Items.FindByValue("SEC");
                if (ddl_cust_cd.Items.Count < 2)
                {
                    ddl_cust_cd.DataSource = dr;
                    ddl_cust_cd.DataValueField = "cust_cd";
                    ddl_cust_cd.DataTextField = "cust_cd";
                    ddl_cust_cd.DataBind();
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        protected void btn_Add_Click(object sender, EventArgs e)
        {
            if (tb_rowcnt.Text == null || tb_rowcnt.Text == "")
            {
                MessageBox.ShowMessage("추가할 Row수를 입력해주세요.", this.Page);
                tb_rowcnt.Focus();
                return;
            }
            else
            {
                FpSpread1_ITEMGR.Sheets[0].AddRows(FpSpread1_ITEMGR.Sheets[0].RowCount, Convert.ToInt16(tb_rowcnt.Text));

            }
        }

        protected void btn_Delete_Click(object sender, EventArgs e)
        {
            System.Collections.IEnumerator enu = FpSpread1_ITEMGR.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;

            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread1_ITEMGR.Sheets[0].ActiveRow;
                //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                for (int i = 0; i < cr.RowCount; i++)
                {
                    string L_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 0].Text;
                    string S_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1].Text;
                    string PROCESS_TYPE = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 2].Text;

                    string sql = "delete T_PROCESSTYPE_GROUP ";
                    sql = sql + " where L_GROUP  ='" + L_GROUP + "' AND S_GROUP  ='" + S_GROUP + "' AND PROCESS_TYPE  ='" + PROCESS_TYPE + "' ";

                    if (QueryExecute(conn_if, sql, "") > 0)
                        FpSpread1_ITEMGR.Sheets[0].Rows.Remove(FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1);
                }
            }
            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
        }

        protected void btn_save_Click(object sender, EventArgs e)
        {
            FpSpread1_ITEMGR.SaveChanges();
            MessageBox.ShowMessage("저장되었습니다.", this.Page);
        }

        protected void btn_exe_Click(object sender, EventArgs e)
        {
            sql = "select L_GROUP,S_GROUP, PROCESS_TYPE from T_PROCESSTYPE_GROUP ";

            sqlAdapter1 = new OracleDataAdapter(sql, conn_if);

            sqlAdapter1.Fill(ds, "ds");

            FpSpread1_ITEMGR.DataSource = ds;
            FpSpread1_ITEMGR.DataBind();
            MessageBox.ShowMessage("조회되었습니다.", this.Page);
        }

        public int QueryExecute(OracleConnection connection, string sql, string wk_type)
        {

            connection.Open();
            cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                //삭제시 기존 권한아이디에 프로그램이 연결되었는지 확인하기 위함.
                if (wk_type == "check")
                    value = Convert.ToInt32(cmd.ExecuteScalar());
                else
                    value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                value = -1;
            }

            connection.Close();
            return value;
        }

        protected void FpSpread1_ITEMGR_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            int colcnt;
            int i;
            int r = (int)e.CommandArgument;
            colcnt = e.EditValues.Count - 1;



            for (i = 0; i <= colcnt; i++)
            {
                if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
                {
                    string sql;

                    //업데이트시
                    if (FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value != null)
                    {
                        /*기존값 가져오기*/
                        string L_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value.ToString();
                        string S_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 1].Value.ToString();
                        string PROCESS_TYPE = FpSpread1_ITEMGR.Sheets[0].Cells[r, 2].Value.ToString();
                        string cg_L_GROUP, cg_S_GROUP, cg_PROCESS_TYPE;
                        /*변경된값 가져오기*/
                        if (i == 0)
                            cg_L_GROUP = e.EditValues[0].ToString();
                        else
                            cg_L_GROUP = L_GROUP;
                        if (i == 1)
                            cg_S_GROUP = e.EditValues[1].ToString();
                        else
                            cg_S_GROUP = S_GROUP;
                        if (i == 2)
                            cg_PROCESS_TYPE = e.EditValues[2].ToString();
                        else
                            cg_PROCESS_TYPE = L_GROUP;

                        
                        sql = "update T_PROCESSTYPE_GROUP ";
                        sql = sql + "set L_GROUP = '" + cg_L_GROUP + "',S_GROUP = '" + cg_S_GROUP + "', PROCESS_TYPE = '" + cg_PROCESS_TYPE + "' ,updt_dt = sysdate ";
                        sql = sql + " where L_GROUP = '" + L_GROUP + "'  ";
                        QueryExecute(conn_if, sql, "");
                    }
                    else
                    {
                        //r = r + 1;
                        //int j = FpSpread1.Sheets[0].ColumnCount;
                        string L_GROUP = e.EditValues[0].ToString();
                        string S_GROUP = e.EditValues[1].ToString();
                        string PROCESS_TYPE = e.EditValues[2].ToString();

                        if (L_GROUP == null || L_GROUP == "" || L_GROUP == "System.Object")
                            MessageBox.ShowMessage("대분류를 입력해주세요.", this.Page);
                        else if (S_GROUP == null || S_GROUP == "" || S_GROUP == "System.Object")
                            MessageBox.ShowMessage("소분류를 입력해주세요.", this.Page);
                        else if (PROCESS_TYPE == null || PROCESS_TYPE == "" || PROCESS_TYPE == "System.Object")
                            MessageBox.ShowMessage("프로세스타입을 입력해주세요.", this.Page);
                        else
                        {
                            sql = "insert into T_PROCESSTYPE_GROUP ";
                            sql = sql + "values('" + L_GROUP + "','" + S_GROUP + "','" + PROCESS_TYPE + "', 'unierp', sysdate, 'unierp', sysdate)";
                            QueryExecute(conn_if, sql, "");

                        }
                    }
                }
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.StoredProcedure;
            sql_cmd.CommandText = "dbo.usp_sale_trand_viewer";
            sql_cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@yyyy", SqlDbType.VarChar, 4);
            SqlParameter param2 = new SqlParameter("@bp_nm", SqlDbType.VarChar, 30);
            
            param1.Value = tb_yyyy.Text;
            param2.Value = ddl_cust_cd.SelectedValue;
            if (tb_yyyy.Text == "" || tb_yyyy.Text == null)
                MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
            else
            {

                sql_cmd.Parameters.Add(param1);
                sql_cmd.Parameters.Add(param2);

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    string report_nm;
                    report_nm = "sm_s2001.rdlc";

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                    ReportViewer1.LocalReport.DisplayName = tb_yyyy.Text + "_매출트랜드_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dt;
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    ReportViewer1.LocalReport.Refresh();
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
            }
        }

        protected void ddl_pg_gubun_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddl_pg_gubun.SelectedValue == "PG_B")
            {
                Panel_default.Visible = false;
                Panel_Default_Btn.Visible = false;
                Panel_Spread_bas.Visible = true;
                Panel_Spread_Btn.Visible = true;
            }
            else
            {
                Panel_default.Visible = true;
                Panel_Default_Btn.Visible = true;
                Panel_Spread_bas.Visible = false;
                Panel_Spread_Btn.Visible = false;
            }
        }

       
    }
}