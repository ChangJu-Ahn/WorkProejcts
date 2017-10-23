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
namespace ERPAppAddition.ERPAddition.MM.MM_M3001
{
    public partial class MM_M3001 : System.Web.UI.Page
    {

        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        //SqlDataAdapter erp_sqlAdapter;
        DataSet ds = new DataSet();

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;
        //string sql_spread;
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

                //MessageBox.ShowMessage(userid, this.Page);

                Session["User"] = userid;
                //rbtn_work_type_SelectedIndexChanged(null, null);

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


        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (rbl_view_type.SelectedValue == "A") //FCST 등록 선택
            {
                panel_upload.Visible = true;
                Panel_select.Visible = false;
                Panel_regist_excel_grid.Visible = true;
                
            }

            if (rbl_view_type.SelectedValue == "B")//FCST 조회 선택
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
            if (list_regist_version.SelectedValue.ToString() == "-선택안함-" || list_regist_version.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("Version을선택해주세요.", this.Page);
                return;
            }


            //}

            if (txt_regist_date_yyyy.Text == "" || txt_regist_date_yyyy.Text == null)
            {
                MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
                return;
            }

            if (txt_regist_date_mm.Text == "" || txt_regist_date_mm.Text == null)
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

        /********************************* 저장 ************************************/


        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
          
            for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
            {
                chk_save_yn = 0;

                string version_no = list_regist_version.SelectedValue.ToString();//version
                string cust_nm = grid_regist_excel.Rows[i].Cells[0].Text.Trim();//거래처
                string item_nm = grid_regist_excel.Rows[i].Cells[1].Text.Trim();//품목명
                string size = grid_regist_excel.Rows[i].Cells[2].Text.Trim();//size
                string process_type = grid_regist_excel.Rows[i].Cells[3].Text.Trim();//process_type
                string route = grid_regist_excel.Rows[i].Cells[4].Text.Trim();//route
                string pkg_type = grid_regist_excel.Rows[i].Cells[5].Text.Trim();//pkg　type 
                string base_mm = txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString(); //기준년월(화면에서 년도선택, 월선택)
                string base_qty = grid_regist_excel.Rows[i].Cells[6].Text.Trim();//당월수량
                string plan_1_qty = grid_regist_excel.Rows[i].Cells[7].Text.Trim();//차월수량
                string plan_2_qty = grid_regist_excel.Rows[i].Cells[8].Text.Trim();//차차월수량
                

                string selectDate = txt_regist_date_mm.SelectedValue.ToString() + "/01/" + txt_regist_date_yyyy.Text;
                DateTime dtTmp = Convert.ToDateTime(selectDate);
                dtTmp = dtTmp.AddMonths(1);

                DateTime dtTmp1 = Convert.ToDateTime(selectDate);
                dtTmp1 = dtTmp.AddMonths(1);

                string sMonth = dtTmp.Month.ToString("00");
                string sYear = dtTmp.Year.ToString();

                string sMonth1 = dtTmp1.Month.ToString("00");
                string sYear1 = dtTmp1.Year.ToString();

                string plan_1_mm = sYear + sMonth;
                string plan_2_mm = sYear1 + sMonth1;

                string sql = "select count(item_nm) from M_FCST_QTY_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'and cust_nm = '" + cust_nm + "' ";
                sql += "and item_nm = '" + item_nm + "' ";
                sql += "and base_mm = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "'";

                if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                {

                    sql = "delete M_FCST_QTY_IMPORT where version_no = '" + list_regist_version.SelectedValue.ToString() + "'"; //and cust_nm = '" + cust_nm + "' ";
                   // sql += " and item_nm = '" + item_nm + "' ";
                   // sql += " and base_qty = '" + base_qty + "'";
                   // sql += " and plan_1_qty = '" + plan_1_qty + "'";
                   // sql += " and plan_2_qty = '" + plan_2_qty + "'";
                    sql += " and base_mm = '" + txt_regist_date_yyyy.Text + txt_regist_date_mm.SelectedValue.ToString() + "'";

                    if (Execute_ERP(conn_erp, sql, "") > 0)
                    {
                        sql = "insert into M_FCST_QTY_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + size + "',";
                        sql += "'" + process_type + "','" + route + "' ,'" + pkg_type + "','0','0','" + base_mm + "',";
                        sql += "'" + plan_1_mm + "','" + plan_2_mm + "',";
                        sql += "'" + base_qty + "','" + plan_1_qty + "','" + plan_2_qty + "','','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                }
                else
                {
                    //엑셀데이타 값이 없는것은 제외:
                    if ((list_regist_version.SelectedValue.ToString() == "" || list_regist_version.SelectedValue.ToString() == null || list_regist_version.SelectedValue.ToString() == "&nbsp;") || (cust_nm == "" || cust_nm == null || cust_nm == "&nbsp;")
                        || (item_nm == "" || item_nm == null || item_nm == "&nbsp;") || (route == "" || route == null || route == "&nbsp;") || (pkg_type == "" || pkg_type == null || pkg_type == "&nbsp;")
                        || (base_mm == "" || base_mm == null || base_mm == "&nbsp;") || (plan_1_mm == "" || plan_1_mm == null || plan_1_mm == "&nbsp;"))
                   
                    {
                        MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                        return;

                    }

                    else
                    {

                        sql = "insert into M_FCST_QTY_IMPORT values('" + list_regist_version.SelectedValue.ToString() + "' ,'" + cust_nm + "' ,'" + item_nm + "','" + size + "',";
                        sql += "'" + process_type + "','" + route + "' ,'" + pkg_type + "','0','0','" + base_mm + "',";
                        sql += "'" + plan_1_mm + "','" + plan_2_mm + "',";
                        sql += "'" + base_qty + "','" + plan_1_qty + "','" + plan_2_qty + "','','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

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
     
            

               
       
      /******************************* 조회 *************************************************************************************/

                      

        protected void btn_select_Click(object sender, EventArgs e)
        {
           
          ReportViewer1.Reset();


          if (txt_select_date_yyyy.Text == "" || txt_select_date_yyyy.Text == null)
          {
              MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
              return;
          }
          if
              (txt_select_date_mm.Text == "" || txt_select_date_mm.Text == null)
          {
              MessageBox.ShowMessage("월을선택해주세요.", this.Page);
              return;
          }

            ReportViewer1.Reset();

            string sql = "SELECT version_no,cust_nm,item_nm,size,process_type,route,pkg_type,base_qty,plan_1_qty,plan_2_qty FROM M_FCST_QTY_IMPORT";
            sql += " Where 1=1 ";

            if (list_select_version.SelectedValue.ToString() != "-선택안함-" && txt_select_date_yyyy.Text != "" && txt_select_date_mm.SelectedValue.ToString() != "-선택안함-") //전부선택
            {
                sql += " and version_no =  '" + list_select_version.SelectedValue.ToString() + "'and base_mm = '" + txt_select_date_yyyy.Text + txt_select_date_mm.Text + "'";
            }

            if (list_select_version.SelectedValue.ToString() != "-선택안함-" && txt_select_date_yyyy.Text != "" && txt_select_date_mm.SelectedValue.ToString() == "-선택안함-")//버전 선택 / 년도 선택 / 월 X
            {
                sql += " and version_no =  '" + list_select_version.SelectedValue.ToString() + "'and base_mm like '" + txt_select_date_yyyy.Text + "%'";
            }

            if (list_select_version.SelectedValue.ToString() != "-선택안함-" && txt_select_date_yyyy.Text == "" && txt_select_date_mm.SelectedValue.ToString() != "-선택안함-")//버전 선택 / 년도 선택X / 월 선택
            {
                MessageBox.ShowMessage("년도를선택해주세요.", this.Page);
                return;
            }

            if (list_select_version.SelectedValue.ToString() != "-선택안함-" && txt_select_date_yyyy.Text == "" ||  txt_select_date_yyyy.Text == null && txt_select_date_mm.SelectedValue.ToString() == "" || txt_select_date_mm.Text == "" || txt_select_date_mm.Text == null) //버전 선택 / 년도 X / 월 X
              {
                  sql += "and version_no =  '" + list_select_version.SelectedValue.ToString() + "'";//선택한 버전에 있는것만
              }
   

            if (list_select_version.SelectedValue.ToString() == "-선택안함-" && txt_select_date_yyyy.Text != "" && txt_select_date_mm.SelectedValue.ToString() != "-선택안함-")//버전 선택X / 년도 선택 / 월 선택
            {
                sql += " and base_mm = '" + txt_select_date_yyyy.Text + txt_select_date_mm.Text + "'";
            }

            if (list_select_version.SelectedValue.ToString() == "-선택안함-" && txt_select_date_yyyy.Text != "" && txt_select_date_mm.SelectedValue.ToString() == "-선택안함-" || txt_select_date_mm.Text == "" || txt_select_date_mm.Text == null)//버전 X / 년도 선택 / 월 X
            {
                sql += "and base_mm like '" + txt_select_date_yyyy.Text + "%'";
            }
       
            sql += " order by cust_nm,item_nm";//고객사, 디바이스 순 정렬

            ReportViewer1.Reset();
            ds_mm_m3001 dt1 = new ds_mm_m3001();

            ReportCreator(dt1, sql, ReportViewer1, "rp_mm_m3001.rdlc", "DataSet1");

        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {

        }

    }
}
