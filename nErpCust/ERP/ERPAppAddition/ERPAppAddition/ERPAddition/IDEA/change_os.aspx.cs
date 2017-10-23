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

namespace ERPAppAddition.ERPAddition.IDEA
{
    public partial class change_os : System.Web.UI.Page
    {

 
        //SqlConnection conn_nact = new SqlConnection(ConfigurationManager.ConnectionStrings["UACT_TEST"].ConnectionString);
        SqlConnection conn_nact = new SqlConnection(ConfigurationManager.ConnectionStrings["UACT"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        
        SqlDataAdapter sqlAdapter;
        
        SqlDataReader dr_nact;

        DataSet ds = new DataSet();

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;


        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {
                ReportViewer1.Reset();
                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_nact.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }


        private void ReportCreator(DataSet _dataSet, string sql, ReportViewer ReportViewer1, string _ReportName, string _ReportDataSourceName)
        {

            conn_nact.Open();
            cmd = conn_nact.CreateCommand();
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = sql;
                dr_nact = cmd.ExecuteReader();
                ds.Tables[0].Load(dr_nact);
                dr_nact.Close();
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
                if (conn_nact.State == ConnectionState.Open)
                    conn_nact.Close();
            }

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
            FpSpread1_ITEMGR.SaveChanges();
            MessageBox.ShowMessage("저장되었습니다.", this.Page);
        }


        //조회
        public void serch_list()
        {
            string sql;

            sql = "SELECT BU_NM,BU_UNICODE,BU_CODE,BU_SERIAL,PRNT_BU_NM,PRNT_DEPT_CD FROM NEPES_OS_STANDARD";

            //sql = "SELECT BU_NM,BU_UNICODE,BU_CODE,BU_SERIAL,PRNT_BU_NM,PRNT_DEPT_CD FROM NEPES_OS_STANDARD";

            sqlAdapter = new SqlDataAdapter(sql, conn_nact);

            sqlAdapter.Fill(ds, "ds");

            FpSpread1_ITEMGR.DataSource = ds;
            FpSpread1_ITEMGR.DataBind();

        }


        protected void btn_exe_Click(object sender, EventArgs e)
        {
            serch_list();
        }

        protected void FpSpread1_ITEMGR_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            int r = (int)e.CommandArgument;
            int colcnt = e.EditValues.Count;

            string sql;

            conn_nact.Open();

            //업데이트시
            if (FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value != null)
            {

                /*기존값가져오기*/
                string BU_NM = FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value.ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
                string BU_UNICODE = FpSpread1_ITEMGR.Sheets[0].Cells[r, 1].Value.ToString();
                string BU_CODE = FpSpread1_ITEMGR.Sheets[0].Cells[r, 2].Value.ToString();
                string BU_SERIAL = FpSpread1_ITEMGR.Sheets[0].Cells[r, 3].Value.ToString();
                string PRNT_BU_NM = FpSpread1_ITEMGR.Sheets[0].Cells[r, 4].Value.ToString();
                string PRNT_DEPT_CD = FpSpread1_ITEMGR.Sheets[0].Cells[r, 5].Value.ToString();

                string cg_BU_NM, cg_BU_UNICODE, cg_BU_CODE, cg_BU_SERIAL, cg_PRNT_BU_NM, cg_PRNT_DEPT_CD;

                /*변경된값가져오기*/
                if (e.EditValues[0].ToString() == "System.Object")
                {
                    cg_BU_NM = BU_NM;
                }
                else
                {
                    cg_BU_NM = e.EditValues[0].ToString();
                }

                if (e.EditValues[1].ToString() == "System.Object")
                {
                    cg_BU_UNICODE = BU_UNICODE;
                }
                else
                {
                    cg_BU_UNICODE = e.EditValues[1].ToString();
                }

                if (e.EditValues[2].ToString() == "System.Object")
                {
                    cg_BU_CODE = BU_CODE;
                }
                else
                {
                    cg_BU_CODE = e.EditValues[2].ToString();
                }

                if (e.EditValues[3].ToString() == "System.Object")
                {
                    cg_BU_SERIAL = BU_SERIAL;
                }
                else
                {
                    cg_BU_SERIAL = e.EditValues[3].ToString();
                }

                if (e.EditValues[4].ToString() == "System.Object")
                {
                    cg_PRNT_BU_NM = PRNT_BU_NM;
                }
                else
                {
                    cg_PRNT_BU_NM = e.EditValues[4].ToString();
                }

                if (e.EditValues[5].ToString() == "System.Object")
                {
                    cg_PRNT_DEPT_CD = PRNT_DEPT_CD;
                }
                else
                {
                    cg_PRNT_DEPT_CD = e.EditValues[5].ToString();
                }

                sql = "UPDATE NEPES_OS_STANDARD SET BU_NM = '" + cg_BU_NM + "',"
                                + "BU_UNICODE = '" + cg_BU_UNICODE + "',"
                                + "BU_CODE = '" + cg_BU_CODE + "',"
                                + "BU_SERIAL = '" + cg_BU_SERIAL + "',"
                                + "PRNT_BU_NM = '" + cg_PRNT_BU_NM + "',"
                                + "PRNT_DEPT_CD = '" + cg_PRNT_DEPT_CD + "',"
                                + "UPDT_USER = 'yoosr',"
                                + "UPDT_DT = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                                + "WHERE BU_NM = '" + BU_NM + "'"
                                + " AND BU_UNICODE = '" + BU_UNICODE + "'"
                                + " AND BU_CODE = '" + BU_CODE + "'"
                                + " AND BU_SERIAL = '" + BU_SERIAL + "'"
                                + " AND PRNT_BU_NM = '" + PRNT_BU_NM + "'"
                                + " AND PRNT_DEPT_CD = '" + PRNT_DEPT_CD + "'";

                SqlCommand sda = new SqlCommand(sql, conn_nact);
                sda.ExecuteNonQuery();


                conn_nact.Close();

                //r = r + 1;
            }
            else
            {

                string BU_NM = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
                string BU_UNICODE = e.EditValues[1].ToString();
                string BU_CODE = e.EditValues[2].ToString();
                string BU_SERIAL = e.EditValues[3].ToString();
                string PRNT_BU_NM = e.EditValues[4].ToString();
                string PRNT_DEPT_CD = e.EditValues[5].ToString();

                if (BU_UNICODE == "System.Object")
                {
                    BU_UNICODE = "";
                }

                sql = "INSERT INTO NEPES_OS_STANDARD(BU_NM,BU_UNICODE,BU_CODE,BU_SERIAL,PRNT_BU_NM,PRNT_DEPT_CD,INSRT_USER,INSRT_DT,UPDT_USER,UPDT_DT) "
                          + "VALUES('" + BU_NM + "','" + BU_UNICODE + "',"
                                       + "'" + BU_CODE + "','" + BU_SERIAL + "',"
                                            + "'" + PRNT_BU_NM + "','" + PRNT_DEPT_CD + "','TEST','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','TEST','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                SqlCommand sda = new SqlCommand(sql, conn_nact);
                sda.ExecuteNonQuery();


                conn_nact.Close();
            }
        }


        protected void btn_Delete_Click(object sender, EventArgs e)
        {

            conn_nact.Open();

            System.Collections.IEnumerator enu = FpSpread1_ITEMGR.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;

            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread1_ITEMGR.Sheets[0].ActiveRow;

                string BU_NM = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 0].Text;
                string BU_UNICODE = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1].Text;
                string BU_CODE = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 2].Text;
                string BU_SERIAL = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 3].Text;
                string PRNT_BU_NM = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 4].Text;
                string PRNT_DEPT_CD = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 5].Text;


                string sql = "DELETE FROM NEPES_OS_STANDARD "
                             + "WHERE BU_NM = '" + BU_NM + "'"
                             + " AND BU_UNICODE = '" + BU_UNICODE + "'"
                             + " AND BU_CODE = '" + BU_CODE + "'"
                             + " AND BU_SERIAL = '" + BU_SERIAL + "'"
                             + " AND PRNT_BU_NM = '" + PRNT_BU_NM + "'"
                             + " AND PRNT_DEPT_CD = '" + PRNT_DEPT_CD + "'";

                SqlCommand sda = new SqlCommand(sql, conn_nact);

                if (sda.ExecuteNonQuery() > 0)
                    FpSpread1_ITEMGR.Sheets[0].Rows.Remove(FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1);
            }

            conn_nact.Close();

            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
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
            InsertOS_People();
        }
        /********************************* 저장 ************************************/

        private void InsertOS_People()
        {
            conn_nact.Open();
            //ADD BY SLYOO : 2014-10-07
            SqlTransaction tran = conn_nact.BeginTransaction();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn_nact;
            cmd.Transaction = tran;
            
            int iComplete = 0;

            for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
            {
                string emp_no = grid_regist_excel.Rows[i].Cells[1].Text.Trim();//사번
                string name = grid_regist_excel.Rows[i].Cells[2].Text.Trim();//성명
                string enter_dt = grid_regist_excel.Rows[i].Cells[4].Text.Trim();//입사일자
                string dept_nm = grid_regist_excel.Rows[i].Cells[5].Text.Trim();//협력사명
                string prnt_dept_nm = grid_regist_excel.Rows[i].Cells[7].Text.Trim();//소속
                string sChange = grid_regist_excel.Rows[i].Cells[11].Text.Trim();//변경점
                string sDate = grid_regist_excel.Rows[i].Cells[12].Text.Trim();//변경일자
                string etc = grid_regist_excel.Rows[i].Cells[13].Text.Trim();//비고

                enter_dt = Convert.ToDateTime(enter_dt).ToString("yyyyMMdd");
                sDate = Convert.ToDateTime(sDate).ToString("yyyyMMdd");

                string sql = "SELECT BU_UNICODE,"
                                       + "BU_CODE+(SELECT MAX(BU_SERIAL)"
                                                    + " FROM NEPES_OS_STANDARD"
                                                    + " WHERE BU_NM LIKE '" + dept_nm + "%'"
                                                    + " AND PRNT_BU_NM LIKE '" + prnt_dept_nm + "%') AS DEPT_CD,"
                                                    + " PRNT_DEPT_CD"
                              + " FROM NEPES_OS_STANDARD"
                                  + " WHERE BU_NM LIKE '" + dept_nm + "%'"
                                  + " AND PRNT_BU_NM LIKE '" + prnt_dept_nm + "%'";



                string BU_UNICODE = "";
                string DEPT_CD = "";
                string PRNT_DEPT_CD = "";

                //SqlCommand sm = new SqlCommand(sql, conn_nact);
                cmd.CommandText = sql;
                SqlDataReader sr = cmd.ExecuteReader();
                

                while (sr.Read())
                {
                    BU_UNICODE = sr.GetString(0);
                    DEPT_CD = sr.GetString(1);
                    PRNT_DEPT_CD = sr.GetString(2);
                }

                sr.Close();

                if (DEPT_CD == "" && sChange == "신규")
                {
                    MessageBox.ShowMessage("저장에 문제가 있습니다. " + name + " 데이타가 기준정보를 벗어나 있습니다.", this);
                    tran.Rollback();
                    break;
                }
                else
                {
                    if (sChange == "신규")
                    {
                        string sql1 = "INSERT INTO dbo.insa_viewer_OS "
                            + "(emp_no,"
                            + "name, "
                            + "dept_cd,"
                            + "emp_position_cd,"
                            + "emp_position, "
                            + "emp_duty_cd, "
                            + "emp_duty, "
                            + "retire_dt,"
                            + "bu_countremark,"
                            + "bu_level,"
                            + "ENTER_DT,"
                            + "dept_nm,"
                            + "prnt_dept_cd,"
                            + "prnt_dept_nm,"
                            + "gubun) "
                                + "values ('" + emp_no + BU_UNICODE + "',"
                                    + "'" + name + "', "
                                    + "'" + DEPT_CD + "',"
                                    + "'228', "
                                    + "'사원(호봉)',"
                                    + "'181', "
                                    + "'부서원',"
                                    + "'',"
                                    + "'000000000000000',"
                                    + "'5',"
                                    + "'" + enter_dt + "',"
                                    + "'" + dept_nm + "',"
                                    + "'" + PRNT_DEPT_CD + "',"
                                    + "'" + prnt_dept_nm + "',"
                                    + "'MP001')";

                        //SqlCommand sda = new SqlCommand(sql1, conn_nact);
                        cmd.CommandText = sql1;
                        cmd.ExecuteNonQuery();

                        iComplete = iComplete + 1;
                    }
                    else if (sChange == "퇴사")
                    {
                        string sql2 = "UPDATE dbo.insa_viewer_OS SET retire_dt = '" + sDate + "' WHERE name = '" + name + "' and dept_nm = '" + dept_nm + "'";

                        //SqlCommand sda = new SqlCommand(sql2, conn_nact);
                        //sda.ExecuteNonQuery();
                        cmd.CommandText = sql2;
                        cmd.ExecuteNonQuery();
                        

                        iComplete = iComplete + 1;
                    }
                    else if (sChange == "전배")
                    {
                        string sql3 = "UPDATE dbo.insa_viewer_OS SET DEPT_NM = '" + etc + "' WHERE name = '" + name + "'and dept_nm = '" + dept_nm + "'";

                        //SqlCommand sda = new SqlCommand(sql3, conn_nact);
                        //sda.ExecuteNonQuery();
                        cmd.CommandText = sql3;
                        cmd.ExecuteNonQuery();

                        iComplete = iComplete + 1;
                    }

                }

            }

            if (grid_regist_excel.Rows.Count == iComplete)
            {
                //실 DB 적용시 주석 풀고 적용
                //string sql4 = "EXEC BATCH_USER_CREATE_STR";
                //cmd.CommandText = sql4;
                //cmd.ExecuteNonQuery();
                
                tran.Commit();
                conn_nact.Close();
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            else
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);
            }

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
            ReportViewer1.Reset();
            
            DataSet_os ch_dt = new DataSet_os();

            string sql4 = "SELECT EMP_NO,NAME,DEPT_NM,DEPT_CD,PRNT_DEPT_NM,PRNT_DEPT_CD,ENTER_DT,RETIRE_DT FROM INSA_VIEWER_OS ";

            if (chk_retire.Checked.ToString() == "True")
            {
                sql4 = sql4 + "WHERE retire_dt <> ''";
            }
            else
            {
                sql4 = sql4 + "WHERE retire_dt = ''";
            }            
            if (txt_emp_no.Text != "")
            {
                sql4 = sql4 + "AND emp_no like '" + txt_emp_no.Text + "%'";
            }
            if (txt_name.Text != "")
            {
                sql4 = sql4 + "AND name like '" + txt_name.Text + "%'";
            }
            if (ddl_dept.Text == "" || ddl_dept.Text == "-선택안함-")
            {
            }
            else
            {
                sql4 = sql4 + "AND dept_nm like '" + ddl_dept.Text + "%'";
            }

            sql4 = sql4 + "order by DEPT_NM";

            ReportCreator(ch_dt, sql4, ReportViewer1, "change_os.rdlc", "DataSet1");           

        }


        protected void ddlSheets_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            ViewExcelSheets(HiddenField_filePath.Value, HiddenField_extension.Value, "Yes");
        }

        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (rbl_view_type.SelectedValue == "A") //기준정보등록 선택
            {

                Panel_Spread_Btn.Visible = true; //스프레드시트 Panel
                panel_upload.Visible = false;
                Panel_Spread_bas.Visible = true;
                //Panel_routeset.Visible = false;
                Panel_select.Visible = false;



            }
            if (rbl_view_type.SelectedValue == "B") //FCST 등록 선택
            {

                Panel_Spread_Btn.Visible = false;
                panel_upload.Visible = true;
                Panel_Spread_bas.Visible = false;
                Panel_select.Visible = false;
                Panel_regist_excel_grid.Visible = true;


            }

            if (rbl_view_type.SelectedValue == "C")//FCST 조회 선택
            {

                Panel_Spread_Btn.Visible = false;
                panel_upload.Visible = false;
                Panel_Spread_bas.Visible = false;
                Panel_select.Visible = true;
                Panel_regist_excel_grid.Visible = false;
                Panel_select_excel_qty_grid.Visible = true;
                //Panel_select_excel_amt_grid.Visible = true;

            }
            ReportViewer1.Reset();         
        }

        protected void btnBackup_Click(object sender, EventArgs e)
        {
            string sql;
            conn_nact.Open();

            sql = "SELECT * INTO dbo.insa_viewer_OS_" + DateTime.Now.ToString("yyyyMMdd") + " FROM dbo.insa_viewer_OS ";

            SqlCommand sda = new SqlCommand(sql, conn_nact);
            sda.ExecuteNonQuery();

            MessageBox.ShowMessage("백업 완료 되었습니다", this);

            conn_nact.Close();

        }
    }
}
