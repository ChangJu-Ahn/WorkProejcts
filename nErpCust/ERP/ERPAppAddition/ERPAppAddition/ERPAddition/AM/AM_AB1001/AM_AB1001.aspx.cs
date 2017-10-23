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
using ERPAppAddition.ERPAddition.SM.sm_sa001;

namespace ERPAppAddition.ERPAddition.AM.AM_AB1001
{
    public partial class AM_AB1001 : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        DataSet ds = new DataSet();
        /*긴급하게 수정 함. 20160302*/
        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleDataReader odr;

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;

        int value, chk_save_yn = 0;
        string userid, db_name;
        cls_dbexe_erp dbexe = new cls_dbexe_erp();

        sa_fun fun = new sa_fun();

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
                /*달력셋*/
                setMonth();
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

        private void setMonth()
        {
            conn.Open();
            try
            {
                DataTable UNIT = fun.getData("SELECT DT FROM T_DAY_INCOM_ROW_INS GROUP BY DT ORDER BY DT DESC");                

                if (UNIT.Rows.Count > 0)
                {
                    ddl_select_date.DataSource = UNIT;
                    ddl_select_date.DataValueField = "DT";
                    ddl_select_date.DataTextField = "DT";
                    ddl_select_date.DataBind();
                    ddl_select_date.SelectedIndex = 0;
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            //return dr;
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

        /********************************* 저장 ************************************/


        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();

            try
            {
                if (RadioButtonPlant.SelectedValue.ToString() == "SEMI")
                {
                    DataTable UNIT = fun.getData("SELECT ISNULL(MAX(NUM), 0) + 1 FROM T_DAY_INCOM_ROW_INS");
                    string NUM = UNIT.Rows[0][0].ToString();

                    for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                    {

                        string DATE = grid_regist_excel.Rows[i].Cells[0].Text.Trim();  //DATE
                        string GUBN = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //구분
                        string PAY_AMT = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //인건비
                        string EQP_AMT = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //감가상각비
                        string ETC_AMT = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //기타경비  

                        sbSQL.Append(" INSERT INTO T_DAY_INCOM_ROW_INS(                           \n");
                        sbSQL.Append(" 	 NUM                                                      \n");
                        sbSQL.Append(" 	,DT                                                       \n");
                        sbSQL.Append(" 	,GUBN                                                     \n");
                        sbSQL.Append(" 	,PAY_AMT                                                  \n");
                        sbSQL.Append(" 	,EQP_AMT                                                  \n");
                        sbSQL.Append(" 	,ETC_AMT                                                  \n");
                        sbSQL.Append(" 	,INSRT_PROG_ID                                            \n");
                        sbSQL.Append(" 	,INSRT_DT                                                 \n");
                        sbSQL.Append(" 	,INSRT_USER_ID                                            \n");
                        sbSQL.Append(" )VALUES(                                                   \n");
                        sbSQL.Append(" 	 '" + NUM + "'                                            \n");
                        sbSQL.Append(" 	,'" + DATE + "'                                           \n");
                        sbSQL.Append(" 	,'" + GUBN.Replace(" ", "").Replace("&quot;", "") + "'      \n");
                        sbSQL.Append(" 	,round('" + PAY_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,round('" + EQP_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,round('" + ETC_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,'AM_AB1001'                                              \n");
                        sbSQL.Append(" 	,GETDATE()                                                \n");
                        sbSQL.Append(" 	,'" + Session["User"].ToString() + " '                        \n");
                        sbSQL.Append(" )                                                          \n");
                    }

                    if (Execute_ERP(conn_erp, sbSQL.ToString(), "") > 0)
                        chk_save_yn += 1;
                }
                else
                {
                    DataTable UNIT = fun.getData("SELECT ISNULL(MAX(NUM), 0) + 1 FROM T_DAY_INCOM_ROW_INS_EM");
                    string NUM = UNIT.Rows[0][0].ToString();

                    for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                    {

                        string DATE = grid_regist_excel.Rows[i].Cells[0].Text.Trim();  //DATE
                        string GUBN = grid_regist_excel.Rows[i].Cells[1].Text.Trim();      //구분
                        string TAF_AMT = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //관세환급
                        string BYP_AMT = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //부산물매각
                        string ETC_AMT = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //기타경비  

                        sbSQL.Append(" INSERT INTO T_DAY_INCOM_ROW_INS_EM(                           \n");
                        sbSQL.Append(" 	 NUM                                                      \n");
                        sbSQL.Append(" 	,DT                                                       \n");
                        sbSQL.Append(" 	,GUBN                                                     \n");
                        sbSQL.Append(" 	,TAF_AMT                                                  \n");
                        sbSQL.Append(" 	,BYP_AMT                                                  \n");
                        sbSQL.Append(" 	,ETC_AMT                                                  \n");
                        sbSQL.Append(" 	,INSRT_PROG_ID                                            \n");
                        sbSQL.Append(" 	,INSRT_DT                                                 \n");
                        sbSQL.Append(" 	,INSRT_USER_ID                                            \n");
                        sbSQL.Append(" )VALUES(                                                   \n");
                        sbSQL.Append(" 	 '" + NUM + "'                                            \n");
                        sbSQL.Append(" 	,'" + DATE + "'                                           \n");
                        sbSQL.Append(" 	,'" + GUBN.Replace(" ", "").Replace("&quot;", "") + "'      \n");
                        sbSQL.Append(" 	,round('" + TAF_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,round('" + BYP_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,round('" + ETC_AMT + "', 0)                                        \n");
                        sbSQL.Append(" 	,'AM_AB1001'                                              \n");
                        sbSQL.Append(" 	,GETDATE()                                                \n");
                        sbSQL.Append(" 	,'" + Session["User"].ToString() + " '                        \n");
                        sbSQL.Append(" )                                                          \n");
                    }

                    if (Execute_ERP(conn_erp, sbSQL.ToString(), "") > 0)
                        chk_save_yn += 1;
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


        /******************************* 조회 *************************************************************************************/



        protected void btn_select_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            try
            {
                string DATE = ddl_select_date.Text;

                StringBuilder sbSQL = new StringBuilder();

                if (RadioButtonPlant.SelectedValue.ToString() == "SEMI")
                {
                    sbSQL.Append(" SELECT DT, GUBN, PAY_AMT, EQP_AMT, ETC_AMT                                  \n");
                    sbSQL.Append(" FROM T_DAY_INCOM_ROW_INS                                                    \n");
                    sbSQL.Append(" WHERE NUM = (SELECT MAX(NUM) FROM T_DAY_INCOM_ROW_INS WHERE DT = '" + DATE + "')\n");
                    sbSQL.Append(" AND DT = '" + DATE + "'                                                         \n");
                    sbSQL.Append(" ORDER BY 2                                                                  \n");
                }
                else
                {
                    sbSQL.Append(" SELECT DT, GUBN, TAF_AMT, BYP_AMT, ETC_AMT                                   \n");
                    sbSQL.Append(" FROM T_DAY_INCOM_ROW_INS_EM                                                    \n");
                    sbSQL.Append(" WHERE NUM = (SELECT MAX(NUM) FROM T_DAY_INCOM_ROW_INS_EM WHERE DT = '" + DATE + "')\n");
                    sbSQL.Append(" AND DT = '" + DATE + "'                                                         \n");
                    sbSQL.Append(" ORDER BY 2                                                                  \n");
                }
                
                DataTable UNIT = fun.getData(sbSQL.ToString());
                UNIT.TableName = "DataSet1";
                
                if(RadioButtonPlant.SelectedValue.ToString() == "SEMI")
                {
                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_ab1001.rdlc");
                }
                else
                {
                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_ab1001_EM.rdlc");
                }
                
                ReportViewer1.LocalReport.DisplayName = "Daily Report 예산등록(NEPES)" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = UNIT;
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
    }
}

          
