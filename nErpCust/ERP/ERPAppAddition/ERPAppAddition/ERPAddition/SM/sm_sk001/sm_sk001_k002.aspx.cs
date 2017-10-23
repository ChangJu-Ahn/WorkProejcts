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

namespace ERPAppAddition.ERPAddition.SM.sm_sk001
{
    public partial class sm_sk001_k002 : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        DataSet ds = new DataSet();
        
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
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select PLAN_YEAR || lpad(PLAN_MONTH, 2, 0) AS M_MONTH   \n");
                sbSQL.Append("  from CALENDAR                                         \n");
                sbSQL.Append(" where PLANT = 'CCUBEDIGITAL'                           \n");
                sbSQL.Append("   and PLAN_YEAR >= '2016'                              \n");
                sbSQL.Append("   and PLAN_YEAR <= to_char(add_months(sysdate, +12), 'yyyy')            \n");
                sbSQL.Append("   and PLAN_YEAR || LPAD(PLAN_MONTH, 2, 0) <= to_char(add_months(sysdate, +12), 'YYYYMM') \n");
                sbSQL.Append(" group by PLAN_YEAR, PLAN_MONTH                         \n");
                sbSQL.Append(" order by 1 desc                                             \n");
                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                odr = cmd2.ExecuteReader();

                if (odr.RowSize > 0)
                {
                    ddl_select_date.DataSource = odr;
                    ddl_select_date.DataValueField = "M_MONTH";
                    ddl_select_date.DataTextField = "M_MONTH";
                    ddl_select_date.DataBind();
                    ddl_select_date.Text = DateTime.Today.AddDays(-1).Year.ToString("0000") + DateTime.Today.AddDays(-1).Month.ToString("00");
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
                DataTable UNIT = fun.getData("SELECT ISNULL(MAX(VER_NO), 0) + 1 FROM T_TEC_IN_USER");
                string VER = UNIT.Rows[0][0].ToString();

                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {

                    string YYYYMM = grid_regist_excel.Rows[i].Cells[0].Text.Trim();       //DATE
                    string PRJ_CD = grid_regist_excel.Rows[i].Cells[1].Text.Trim();       //프로젝트코드
                    string GRP_CD1 = grid_regist_excel.Rows[i].Cells[2].Text.Trim();      //GROUP1
                    string GRP_NM1 = grid_regist_excel.Rows[i].Cells[3].Text.Trim();      //GROUP1
                    string GRP_CD2 = grid_regist_excel.Rows[i].Cells[4].Text.Trim();      //GROUP2
                    string GRP_NM2 = grid_regist_excel.Rows[i].Cells[5].Text.Trim();      //GROUP2
                    string GRP_CD3 = grid_regist_excel.Rows[i].Cells[6].Text.Trim();      //GROUP3
                    string GRP_NM3 = grid_regist_excel.Rows[i].Cells[7].Text.Trim();      //GROUP3
                    string GRP_CD4 = grid_regist_excel.Rows[i].Cells[8].Text.Trim();      //GROUP4
                    string GRP_NM4 = grid_regist_excel.Rows[i].Cells[9].Text.Trim();      //GROUP4
                    string GRP_CD5 = grid_regist_excel.Rows[i].Cells[10].Text.Trim();      //GROUP5
                    string GRP_NM5 = grid_regist_excel.Rows[i].Cells[11].Text.Trim();      //GROUP5
                    string AGB02 = grid_regist_excel.Rows[i].Cells[12].Text.Trim();        //직급코드
                    string EMP_NO = grid_regist_excel.Rows[i].Cells[13].Text.Trim();       //사번 
                    string AMT = grid_regist_excel.Rows[i].Cells[14].Text.Trim();          //AMT
                    
                    sbSQL.Append(" INSERT INTO T_TEC_IN_USER(                                 \n");
                    sbSQL.Append(" 	 VER_NO                                                   \n");
                    sbSQL.Append(" 	,YYYYMM                                                   \n");
                    sbSQL.Append(" 	,PRJ_CD                                                   \n");
                    sbSQL.Append(" 	,GRP_CD1                                                  \n");
                    sbSQL.Append(" 	,GRP_NM1                                                  \n");
                    sbSQL.Append(" 	,GRP_CD2                                                  \n");
                    sbSQL.Append(" 	,GRP_NM2                                                  \n");
                    sbSQL.Append(" 	,GRP_CD3                                                  \n");
                    sbSQL.Append(" 	,GRP_NM3                                                  \n");
                    sbSQL.Append(" 	,GRP_CD4                                                  \n");
                    sbSQL.Append(" 	,GRP_NM4                                                  \n");
                    sbSQL.Append(" 	,GRP_CD5                                                  \n");
                    sbSQL.Append(" 	,GRP_NM5                                                  \n");
                    sbSQL.Append(" 	,AGB02                                                    \n");
                    sbSQL.Append(" 	,EMP_NO                                                   \n");
                    sbSQL.Append(" 	,AMT                                                      \n");
                    sbSQL.Append(" 	,INSRT_PROG_ID                                            \n");
                    sbSQL.Append(" 	,INSRT_DT                                                 \n");                    
                    sbSQL.Append(" )VALUES(                                                   \n");
                    sbSQL.Append(" 	 '" + VER + "'                                            \n");
                    sbSQL.Append(" 	,'" + YYYYMM + "'                                           \n");
                    sbSQL.Append(" 	,'" + PRJ_CD + "'                                           \n");
                    sbSQL.Append(" 	,'" + GRP_CD1.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_NM1.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_CD2.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_NM2.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_CD3.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_NM3.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_CD4.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_NM4.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_CD5.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + GRP_NM5.Replace(" ", "").Replace("&quot;", "").Replace("NULL", "").Replace("&nbsp;", "") + "'      \n");
                    sbSQL.Append(" 	,'" + AGB02 + "'                                         \n");
                    sbSQL.Append(" 	,'" + EMP_NO + "'                                        \n");
                    sbSQL.Append(" 	,'" + AMT + "'                                           \n");
                    sbSQL.Append(" 	,'K002_PRG'                                              \n");
                    sbSQL.Append(" 	,CONVERT(CHAR(8),GETDATE(),112)+REPLACE(CONVERT(CHAR(8),GETDATE(),108),':','')   \n");                    
                    sbSQL.Append(" )                                                         \n");
                }

                if (Execute_ERP(conn_erp, sbSQL.ToString(), "") > 0)
                    chk_save_yn += 1;

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
                string YYYYMM = ddl_select_date.Text;

                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append(" SELECT YYYYMM ,PRJ_CD ,GRP_CD1 ,GRP_NM1 ,GRP_CD2 ,GRP_NM2 ,GRP_CD3 ,GRP_NM3 ,GRP_CD4 ,GRP_NM4 ,GRP_CD5 ,GRP_NM5  \n");
                sbSQL.Append("        ,AGB02, D.UD_MINOR_NM AS AGB02_NM ,U.EMP_NO, T.PR_NAME AS NAME ,AMT                             \n");
                sbSQL.Append(" FROM T_TEC_IN_USER U LEFT JOIN INSADB.inbus.dbo.h_person T ON U.EMP_NO = T.PR_SNO                     \n");
                sbSQL.Append("                      LEFT JOIN B_USER_DEFINED_MINOR D ON U.AGB02 = D.UD_MINOR_CD AND UD_MAJOR_CD = 'AGB02'  \n");
                sbSQL.Append(" WHERE VER_NO = (SELECT MAX(VER_NO) FROM T_TEC_IN_USER WHERE YYYYMM = '" + YYYYMM + "')    \n");
                sbSQL.Append(" AND YYYYMM = '" + YYYYMM + "'                                                         \n");
                sbSQL.Append(" ORDER BY 1, 2                                                                         \n");
                DataTable UNIT = fun.getData(sbSQL.ToString());
                UNIT.TableName = "DataSet1";

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sk001_k002.rdlc");
                ReportViewer1.LocalReport.DisplayName = "인원별배부기준등록" + DateTime.Now.ToShortDateString();
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


