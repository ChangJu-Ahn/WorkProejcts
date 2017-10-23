
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
using Microsoft.Reporting.WebForms;

namespace ERPAppAddition.ERPAddition.AM.AM_A9007
{
    public partial class AM_A9007 : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        DataSet ds = new DataSet();

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;


        FarPoint.Web.Spread.SpreadCommandEventArgs chk;

        int value, chk_save_yn = 0;
        string userid, db_name;
        cls_dbexe_erp dbexe = new cls_dbexe_erp();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "") //사용자 ID값이 없다면 개발자 ID로할지 판단하기
                {
                    if (Request.QueryString["db"] == null || Request.QueryString["db"] == "") //DB없이 바로 실행할때 개발자용으로 적용
                        userid = "dev"; //erp에서 실행하지 않았을시 대비용
                    else // DB명이 있는데 사용자 ID가 없다면 이상하니 다시 접속하라는 메세지 보여줌
                    {
                        MessageBox.ShowMessage("잘못된 접근입니다. ERP접속 후 실행해주세요", this.Page);
                        this.Response.Redirect("../../Fail_Page.aspx");
                    }
                }
                else
                    userid = Request.QueryString["userid"];

                Session["User"] = userid;


                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_date;                
                SqlDataReader dr1;

                setWeek();

                ls_date = "select distinct(dt)  yyyymm from I_WH_COST2 order by dt desc";
                conn.Open();
                SqlCommand cmd2 = new SqlCommand(ls_date, conn);
                
                dr1 = cmd2.ExecuteReader();
                if (txt_select_date_yyyymm.Items.Count < 2)
                {
                    txt_select_date_yyyymm.DataSource = dr1;
                    txt_select_date_yyyymm.DataValueField = "yyyymm";
                    txt_select_date_yyyymm.DataTextField = "yyyymm";
                    txt_select_date_yyyymm.DataBind();
                }
                dr1.Close();

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

        private void setWeek()
        {
            conn.Open();
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select SUBSTR(PLAN_YEAR,3,2)|| '년' || WEEK || '주차(' || NATURAL_DATE ||')'as m_week , NATURAL_DATE  		\n");
                sbSQL.Append("FROM (                                                                                                                      \n");
                sbSQL.Append("    select PLAN_YEAR, NATURAL_DATE, to_char(to_date(NATURAL_DATE,'yyyymmdd'), 'dy') AS DY, LPAD(PLAN_WEEK,2,'0') AS WEEK    \n");
                sbSQL.Append("      from CALENDAR                                                                                                         \n");
                sbSQL.Append("     where PLANT = 'CCUBEDIGITAL'                                                                                           \n");
                sbSQL.Append("       and PLAN_YEAR >= '2015'                                                                                               \n");
                sbSQL.Append("       and PLAN_YEAR <= to_char(sysdate+7, 'yyyy')                                                                            \n");
                sbSQL.Append("       and NATURAL_DATE <= to_char(sysdate+7,'yyyymmdd')                                                                      \n");
                sbSQL.Append("       )                                                                                                                    \n");
                sbSQL.Append("WHERE DY = '금'                                                                                                             \n");
                sbSQL.Append("order by NATURAL_DATE  desc                                                                                                     \n");

                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                dr = cmd2.ExecuteReader();

                if (dr.RowSize > 0)
                {
                    txt_regist_date_yyyymm.DataSource = dr;
                    txt_regist_date_yyyymm.DataValueField = "NATURAL_DATE";
                    txt_regist_date_yyyymm.DataTextField = "m_week";
                    txt_regist_date_yyyymm.DataBind();
                    txt_regist_date_yyyymm.SelectedIndex = 0;
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
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


        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e) // 구분 선택
        {

            if (rbl_view_type.SelectedValue == "A") //예산 등록 선택
            {
                panel_upload.Visible = true;
                Panel_select.Visible = false;
                Panel_regist_excel_grid.Visible = true;

            }

            if (rbl_view_type.SelectedValue == "B")//예산 조회 선택
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_month_sql;
                string ls_date;
                SqlDataReader dr;
                SqlDataReader dr1;

                ls_month_sql = "SELECT top 1 convert(varchar(8),getdate(),112) yyyymm from B_ITEM_BY_PLANT (nolock) where ITEM_CD ='1CAB-08002-2'";
                SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);


                ls_date = "select distinct(dt)  yyyymm from I_WH_COST2 order by dt desc";


                SqlCommand cmd2 = new SqlCommand(ls_date, conn);

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;


                cmd = conn.CreateCommand();

                dr = cmd5.ExecuteReader();


                dr.Close();

                dr1 = cmd2.ExecuteReader();

                txt_select_date_yyyymm.DataSource = dr1;
                txt_select_date_yyyymm.DataValueField = "yyyymm";
                txt_select_date_yyyymm.DataTextField = "yyyymm";
                txt_select_date_yyyymm.DataBind();

                dr1.Close();
                panel_upload.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;


            }



            if (rbl_view_type.SelectedValue == "C")//예산 조회 선택
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_month_sql;
                string ls_date;
                SqlDataReader dr;
                SqlDataReader dr1;

                ls_month_sql = "SELECT top 1 convert(varchar(8),getdate(),112) yyyymm from B_ITEM_BY_PLANT (nolock) where ITEM_CD ='1CAB-08002-2'";
                SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);


                ls_date = "select distinct(INSPEC_DT)  yyyymm from EM_SOTCK_INSPEC order by INSPEC_DT desc";


                SqlCommand cmd2 = new SqlCommand(ls_date, conn);


                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;


                cmd = conn.CreateCommand();

                dr = cmd5.ExecuteReader();



                dr.Close();

                dr1 = cmd2.ExecuteReader();

                txt_select_date_yyyymm.DataSource = dr1;
                txt_select_date_yyyymm.DataValueField = "yyyymm";
                txt_select_date_yyyymm.DataTextField = "yyyymm";
                txt_select_date_yyyymm.DataBind();

                dr1.Close();
                panel_upload.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;

            }


            if (rbl_view_type.SelectedValue == "D")//예산 조회 선택
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_month_sql;
                string ls_date;
                SqlDataReader dr;
                SqlDataReader dr1;

                ls_month_sql = "SELECT top 1 convert(varchar(8),getdate(),112) yyyymm from B_ITEM_BY_PLANT (nolock) where ITEM_CD ='1CAB-08002-2'";
                SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);


                ls_date = "select distinct(INSPEC_DT)  yyyymm from DS_SOTCK_INSPEC order by INSPEC_DT desc";


                SqlCommand cmd2 = new SqlCommand(ls_date, conn);

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;


                cmd = conn.CreateCommand();

                dr = cmd5.ExecuteReader();


                dr.Close();

                dr1 = cmd2.ExecuteReader();

                txt_select_date_yyyymm.DataSource = dr1;
                txt_select_date_yyyymm.DataValueField = "yyyymm";
                txt_select_date_yyyymm.DataTextField = "yyyymm";
                txt_select_date_yyyymm.DataBind();
        

                dr1.Close();
                panel_upload.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;

            }


            if (rbl_view_type.SelectedValue == "E")//예산 조회 선택
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_month_sql;
                string ls_date;
                SqlDataReader dr;
                SqlDataReader dr1;

                ls_month_sql = "SELECT top 1 convert(varchar(8),getdate(),112) yyyymm from EM_DAY_GOODS (nolock) where ITEM_CD ='1CAB-08002-2'";
                SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);


                ls_date = "select distinct(YYYYMM)  yyyymm from EM_DAY_GOODS order by YYYYMM desc";


                SqlCommand cmd2 = new SqlCommand(ls_date, conn);

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;


                cmd = conn.CreateCommand();

                dr = cmd5.ExecuteReader();


                dr.Close();

                dr1 = cmd2.ExecuteReader();

                txt_select_date_yyyymm.DataSource = dr1;
                txt_select_date_yyyymm.DataValueField = "yyyymm";
                txt_select_date_yyyymm.DataTextField = "yyyymm";
                txt_select_date_yyyymm.DataBind();


                dr1.Close();
                panel_upload.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;

            }


            if (rbl_view_type.SelectedValue == "F")//예산 조회 선택
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlCommand cmd1 = new SqlCommand();
                SqlDataReader dReader_select;
                string id = "";
                string ls_biz_area_cd_sql;
                string ls_month_sql;
                string ls_date;
                SqlDataReader dr;
                SqlDataReader dr1;

                ls_month_sql = "SELECT top 1 convert(varchar(8),getdate(),112) yyyymm from EM_DAY_GOODS (nolock) where ITEM_CD ='1CAB-08002-2'";
                SqlCommand cmd5 = new SqlCommand(ls_month_sql, conn);


                ls_date = "select top 1 convert(varchar(8),getdate(),112) yyyymm  from EM_DAY_GOODS ";


                SqlCommand cmd2 = new SqlCommand(ls_date, conn);

                conn.Open();
                cmd = conn.CreateCommand();
                cmd.CommandTimeout = 500;
                cmd.CommandType = CommandType.Text;


                cmd = conn.CreateCommand();

                dr = cmd5.ExecuteReader();


                dr.Close();

                dr1 = cmd2.ExecuteReader();

                txt_select_date_yyyymm.DataSource = dr1;
                txt_select_date_yyyymm.DataValueField = "yyyymm";
                txt_select_date_yyyymm.DataTextField = "yyyymm";
                txt_select_date_yyyymm.DataBind();


                dr1.Close();
                panel_upload.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;

            }



            ReportViewer1.Reset();

        }
        /*********************************엑셀업로드 ************************************/

        // 엑셀업로드 클릭
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            //if (ddl_plan_version.SelectedValue.ToString() == "-선택안함-" || ddl_plan_version.SelectedValue.ToString() == null)
            //{
            //    MessageBox.ShowMessage("계획 Version을선택해주세요.", this.Page);
            //    return;
            //}


            if (txt_regist_date_yyyymm.Text == "" || txt_regist_date_yyyymm.Text == null)
            {
                MessageBox.ShowMessage("년월을 입력해주세요.", this.Page);
                return;
            }

            //if (ddl_biz_cd.SelectedValue.ToString() == "-선택안함-" || ddl_biz_cd.SelectedValue.ToString() == null)
            //{
            //    MessageBox.ShowMessage("사업장을 입력해주세요.", this.Page);
            //    return;
            //}


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

        /********************************* 저장 ************************************/


        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            chk_save_yn = 0;
            try
            {   
                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {
                    string yyyymm = txt_regist_date_yyyymm.Text.ToString();           //년월
                    //string plan_flg = ddl_plan_version.SelectedValue.ToString();      //계획버전
                    //string biz_cd = ddl_biz_cd.SelectedValue.ToString();             //사업장
                    //string acct_cd = grid_regist_excel.Rows[i].Cells[0].Text.Trim();  //계정코드
                    //string amt = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //금액


                    string fac = grid_regist_excel.Rows[i].Cells[0].Text.Trim();      //공장
                    string wh = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //창고
                    string item_cd = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //ITEM_CD
                    string item_nm = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //ITEM_NM
                    string unit = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //UNIT
                    string qty = grid_regist_excel.Rows[i].Cells[5].Text.Trim();      //QTY

                    /*비고란 입력 여부 확인*/
                    string remark = "";
                    if (grid_regist_excel.Rows[0].Cells.Count > 6)
                    {
                        remark = grid_regist_excel.Rows[i].Cells[6].Text.Trim() == "&nbsp;" ? "" : grid_regist_excel.Rows[i].Cells[6].Text.Trim();      //REMARK
                    }
                    
                    string sql = "select COUNT(item_cd)  from i_wh_cost2 where dt = '" + yyyymm + "'";

                   if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {
                        if (i == 0)
                        {
                            sql = "delete i_wh_cost2 where dt = '" + yyyymm + "'";
                            Execute_ERP(conn_erp, sql, "");
                        }
                            sql = "insert into i_wh_cost2 values('" + yyyymm + "' ,'" + fac + "' ,'" + wh + "','" + item_cd + "','" + item_nm + "','" + unit + "','" + qty + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', '" + remark + "' )";
                          
                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ( (yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;"))
                            //if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;") || (fac == "" || fac == null || fac == "&nbsp;") || (wh == "" || wh == null || wh == "&nbsp;") ||
                            // (item_cd == "" || item_cd == null || item_cd == "&nbsp;") || (item_nm == "" || item_nm == null || item_nm == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;

                        }

                        else
                        {
                            sql = "insert into i_wh_cost2 values('" + yyyymm + "' ,'" + fac + "' ,'" + wh + "','" + item_cd + "','" + item_nm + "','" + unit + "','" + qty + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','"+ remark+ "' )";

                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                }
                MessageBox.ShowMessage("["+ chk_save_yn + " ] 의 값 저장되었습니다", this);
            }
            catch
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);

            }
            ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }



        /******************************* 조회 *************************************************************************************/



        protected void btn_select_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.USP_I_WH_COST";
            //cmd_erp.CommandText = "dbo.USP_I_WH_COST_2";
            cmd_erp.CommandTimeout = 3000;

  
            SqlParameter param1 = new SqlParameter("@S_I_DT", SqlDbType.VarChar, 8);
            SqlParameter param2 = new SqlParameter("@r_flag", SqlDbType.VarChar, 1);

            //SqlParameter param1 = new SqlParameter("@S_I_DT", SqlDbType.VarChar, 8);
            //SqlParameter param2 = new SqlParameter("@PLAN_FLG", SqlDbType.VarChar, 4);
            //SqlParameter param3 = new SqlParameter("@BIZ_CD", SqlDbType.VarChar, 4);


            string YYYYMM, PLAN_FLG, BIZ_CD;
            YYYYMM = txt_select_date_yyyymm.Text;
            //PLAN_FLG = ddl_select_version.SelectedValue;
            //BIZ_CD = ddl_select_biz.SelectedValue;

            param1.Value = YYYYMM;
            if (YYYYMM == null || YYYYMM == "")
                YYYYMM = "%";

            if (rbl_view_type.SelectedValue == "B")
               {param2.Value ='1' ;}
            else if (rbl_view_type.SelectedValue == "C")
               {param2.Value = '2'  ; }
            else if (rbl_view_type.SelectedValue == "D")
               { param2.Value = '3'; }
            else if (rbl_view_type.SelectedValue == "E")
               { param2.Value = '4'; }
            else
               { param2.Value = '5'; }


            //param2.Value = PLAN_FLG;
            //if (PLAN_FLG == null || PLAN_FLG == "" || PLAN_FLG == "0")
            //    PLAN_FLG = "%";

            //param3.Value = BIZ_CD;
            //if (BIZ_CD == null || BIZ_CD == "")
            //    BIZ_CD = "%";


            cmd_erp.Parameters.Add(param1);
            cmd_erp.Parameters.Add(param2);
            //cmd_erp.Parameters.Add(param3);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_a9007.rdlc");
                ReportViewer1.LocalReport.DisplayName = "현장재고 금액 조회" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

                //UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
        }

        protected void txt_regist_date_yyyymm_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void btnCancel0_Click(object sender, EventArgs e)
        {

            ////전사재고 집계실행
            //string ls_sum_dt = txt_select_date_yyyymm.Text;

            //string sql = " exec USP_TOTAL_STOCK_BATCH  '" + ls_sum_dt + "'";


            //if (Execute_ERP(conn_erp, sql, "") < 0 )
            //{
            //    MessageBox.ShowMessage("오류 데이터가 있습니다.", this.Page);
            //    return;
            //}

            //    MessageBox.ShowMessage("저장 되었습니다.", this.Page);


        }

        protected void btnemSave_Click(object sender, EventArgs e)
        {
            chk_save_yn = 0;
            try
            {   
                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {
                    string yyyymm = txt_regist_date_yyyymm.Text.ToString();           //년월
                    string m_gubn = grid_regist_excel.Rows[i].Cells[0].Text.Trim();   //자재구분
                    string item_nm = grid_regist_excel.Rows[i].Cells[1].Text.Trim();  //ITEM_NM
                    string qty = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //QTY
                    string price = grid_regist_excel.Rows[i].Cells[3].Text.Trim();    //price
                    string amt = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //amt

                    /*원자재 부자재 재공 나누어서 입력하는지? 의문이 든다.*/
                    string ls_checkstring = yyyymm + m_gubn + item_nm;

                    string sql = "select COUNT(item_nm)  from EM_SOTCK_INSPEC where inspec_dt+m_gubn+item_nm = '" + ls_checkstring + "'";

                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {
                        sql = "delete EM_SOTCK_INSPEC where inspec_dt+m_gubn+item_nm = '" + ls_checkstring + "'";
                        Execute_ERP(conn_erp, sql, "");

                        sql = "insert into EM_SOTCK_INSPEC values('" + yyyymm + "' ,'" + m_gubn + "' ,'" + item_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;
                        }
                        else
                        {
                            sql = "insert into EM_SOTCK_INSPEC values('" + yyyymm + "' ,'" + m_gubn + "' ,'" + item_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";
                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                }
                MessageBox.ShowMessage("[" + chk_save_yn + " ] 의 값 저장되었습니다", this);
            }
            catch
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);
            }
            ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }

        protected void btndsSave_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {
                    chk_save_yn = 0;

                    string yyyymm = txt_regist_date_yyyymm.Text.ToString();           //년월
                    string m_gubn = grid_regist_excel.Rows[i].Cells[0].Text.Trim();      //자재구분
                    string item_nm = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //ITEM_NM
                    string qty = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //QTY
                    string price = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //price
                    string amt = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //amt

                    string sql = "select COUNT(item_nm)  from DS_SOTCK_INSPEC where inspec_dt = '" + yyyymm + "'";

                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {
                        if (i == 0)
                        {
                            sql = "delete DS_SOTCK_INSPEC where INSPEC_DT = '" + yyyymm + "'";
                            Execute_ERP(conn_erp, sql, "");
                        }

                        sql = "insert into DS_SOTCK_INSPEC values('" + yyyymm + "' ,'" + m_gubn + "' ,'" + item_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        //if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;") || (m_gubn == "" || m_gubn == null || m_gubn == "&nbsp;") || (item_nm == "" || item_nm == null || item_nm == "&nbsp;"))
                        if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;
                        }

                        else
                        {
                            sql = "insert into DS_SOTCK_INSPEC values('" + yyyymm + "' ,'" + m_gubn + "' ,'" + item_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                }
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            catch
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);

            }
            ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }

        protected void btnemgoodsSave_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {
                    chk_save_yn = 0;

                    //string yyyymm = txt_regist_date_yyyymm.Text.ToString();           //년월
                    string yyyymm = grid_regist_excel.Rows[i].Cells[0].Text.Trim();           //년월
                    string dt = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //매출일
                    string item_cd = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //상품코드
                    string cust_nm = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //매출처
                    string qty = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //QTY
                    string price = grid_regist_excel.Rows[i].Cells[5].Text.Trim();      //price
                    string amt = grid_regist_excel.Rows[i].Cells[6].Text.Trim();      //amt
                    string gubn = grid_regist_excel.Rows[i].Cells[7].Text.Trim();      //제품구분

                    string ls_checkstring = yyyymm ;
                    //string sql;

                    string sql = "select COUNT(item_cd)  from EM_DAY_GOODS where YYYYMM= '" + ls_checkstring + "'";

                    //if (i == 0)
                    //{
                    //    sql = "delete EM_DAY_GOODS where YYYYMM = '" + ls_checkstring + "'";
                    //    Execute_ERP(conn_erp, sql, "");
                    //}


                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {
                         if (i == 0)
                         {
                        sql = "delete EM_DAY_GOODS where YYYYMM = '" + ls_checkstring + "'";

                        //sql = "delete EM_DAY_GOODS where item_cd = '" + ls_checkstring + "'";

                        Execute_ERP(conn_erp, sql, "");
                          }

                         sql = "insert into EM_DAY_GOODS values('" + yyyymm + "' ,'" + dt + "' ,'" + item_cd + "','" + cust_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + gubn + "' )";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;"))
                            //if ((yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;") || (dt == "" || dt == null || dt == "&nbsp;") || (item_cd == "" || item_cd == null || item_cd == "&nbsp;") || (cust_nm == "" || cust_nm == null || cust_nm == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;
                        }

                        else
                        {

                            sql = "insert into EM_DAY_GOODS values('" + yyyymm + "' ,'" + dt + "' ,'" + item_cd + "','" + cust_nm + "','" + qty + "','" + price + "','" + amt + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + gubn + "'  )";

                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                }
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            catch
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);

            }
            ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }



        protected void btn_exec_Click(object sender, EventArgs e)
        {
            if(rbl_view_type.SelectedValue == "E")
            {
                MessageBox.ShowMessage("EM상품매출 조회에서는 실행할 수 없습니다.", this.Page);
                return;
            }
            //전사재고 집계실행
            string ls_sum_dt = txt_select_date_yyyymm.Text;

            string sql = " exec USP_TOTAL_STOCK_BATCH  '" + ls_sum_dt + "'";


            if (Execute_ERP(conn_erp, sql, "") < 0)
            {
                MessageBox.ShowMessage("오류 데이터가 있습니다.", this.Page);
                return;
            }

            MessageBox.ShowMessage("저장 되었습니다.", this.Page);

        }
    }
}


