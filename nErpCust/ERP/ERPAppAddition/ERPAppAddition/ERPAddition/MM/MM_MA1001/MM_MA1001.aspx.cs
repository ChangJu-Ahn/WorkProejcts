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
using FarPoint.Web.Spread.Model;

namespace ERPAppAddition.ERPAddition.MM.MM_MA1001
{
    public partial class MM_MA1001 : System.Web.UI.Page
    {
        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        DataSet ds = new DataSet();
        SqlDataAdapter sqlAdapter1;

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

        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e) // 구분 선택
        {

            if (rbl_view_type.SelectedValue == "A") //신규 등록 선택
            {
                panel_upload.Visible = true;
                Panel_regist_excel_grid.Visible = true;
                Panel_menu.Visible = false;
                Panel_del.Visible = false;
                panel_spread.Visible = false;
                    

            }

            if (rbl_view_type.SelectedValue == "B")//조회/수정/삭제 선택
            {
                panel_upload.Visible = false;
                Panel_regist_excel_grid.Visible = false;
                Panel_menu.Visible = true;
                Panel_del.Visible = true;
                panel_spread.Visible = true;


            }

        }
        /*********************************엑셀업로드 ************************************/

        // 엑셀업로드 클릭
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (ddl_item_cd.SelectedValue.ToString() == "-선택안함-" || ddl_item_cd.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("공장을선택해주세요.", this.Page);
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

        /********************************* 엑셀업로드 ************************************/


        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {

                for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                {
                    chk_save_yn = 0;

                                 
                    string erp_item_cd = ddl_item_cd.Text.ToString();                    //ERP품목코드
                    string item_nm = grid_regist_excel.Rows[i].Cells[0].Text.Trim();         //품목명
                    string item_group1 = grid_regist_excel.Rows[i].Cells[1].Text.Trim();     //대분류
                    string item_group2 = grid_regist_excel.Rows[i].Cells[2].Text.Trim();     //중분류
                    string spec = grid_regist_excel.Rows[i].Cells[3].Text.Trim();            //SPEC
                    string unit = grid_regist_excel.Rows[i].Cells[4].Text.Trim();            //UNIT
                    string amt = grid_regist_excel.Rows[i].Cells[5].Text.Trim();             //AMT
                    string moq = grid_regist_excel.Rows[i].Cells[6].Text.Trim();             //MOQ
                    string item_day = grid_regist_excel.Rows[i].Cells[7].Text.Trim();        //ITEM_DAY
                    string mro_item_cd = grid_regist_excel.Rows[i].Cells[8].Text.Trim();    //MRO_ITEM_CD



                    string sql = "select COUNT(mro_item_cd) from b_item_mro where item_cd = '" + erp_item_cd + "' and mro_item_cd = '" + mro_item_cd + "' ";
                   
                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {

                        sql = "delete b_item_mro where item_cd_ = '" + erp_item_cd + "' and mro_item_cd = '" + mro_item_cd + "' ";
                        

                        if (Execute_ERP(conn_erp, sql, "") > 0)
                        {
                            sql = "insert into b_item_mro values('" + erp_item_cd + "' ,'" + item_nm + "' ,'" + item_group1 + "','" + item_group2 + "',";
                            sql = "'" + spec + "' ,'" + unit + "' ,'" + amt + "','" + moq + "','" + item_day + "','" + mro_item_cd + "','Y',";
                            sql += "'" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";

                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((erp_item_cd == "" || erp_item_cd == null || erp_item_cd == "&nbsp;") || (item_nm == "" || item_nm == null || item_nm == "&nbsp;")
                            || (mro_item_cd == "" || mro_item_cd == null || mro_item_cd == "&nbsp;") )
                        {
                            MessageBox.ShowMessage(i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);
                            return;

                        }

                        else
                        {
                            sql = "insert into b_item_mro values('" + erp_item_cd + "' ,'" + item_nm + "' ,'" + item_group1 + "','" + item_group2 + "',";
                            sql += "'" + spec + "' ,'" + unit + "' ,'" + amt + "','" + moq + "','" + item_day + "','" + mro_item_cd + "','Y',";
                            sql += "'" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";


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

            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }
            protected void FpSpread1_ActiveRowChanged(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {

        }

            /********************************* 저장 ************************************/
            protected void btn_save_Click(object sender, EventArgs e)
            {
                FpSpread1.SaveChanges();
                //FpSpread1.Reset();
                MessageBox.ShowMessage("저장되었습니다.", this.Page);
                btn_search_Click(this, new EventArgs());
            }


            protected void FpSpread1_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
            {

                int colcnt;
                int i;
                string cg_mro_item_cd, cg_item_group1, cg_item_group2, cg_item_nm, cg_spec, cg_unit, cg_amt, cg_moq, cg_item_day, cg_valid_flg ;
               
                int r = (int)e.CommandArgument;
                colcnt = e.EditValues.Count - 1;


                for (i = 0; i <= colcnt; i++)
                {
                    if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
                    {
                        string sql;

                        /*기존값 가져오기*/

                        string erp_item_cd = ddl_item_cd2.SelectedValue.ToString();              //ERP품목코드
                        string mro_item_cd = FpSpread1.Sheets[0].Cells[r, 0].Value.ToString();
                        string item_group1 = FpSpread1.Sheets[0].Cells[r, 1].Value.ToString();  //대분류
                        string item_group2 = FpSpread1.Sheets[0].Cells[r, 2].Value.ToString();  //중분류
                        string item_nm     = FpSpread1.Sheets[0].Cells[r, 3].Value.ToString();  //품목명
                        string spec = FpSpread1.Sheets[0].Cells[r, 4].Value.ToString();         //SPEC
                        string unit = FpSpread1.Sheets[0].Cells[r, 5].Value.ToString();         //UNIT
                        string amt = FpSpread1.Sheets[0].Cells[r, 6].Value.ToString();          //AMT
                        string moq = FpSpread1.Sheets[0].Cells[r, 7].Value.ToString();          //MOQ
                        string item_day = FpSpread1.Sheets[0].Cells[r, 8].Value.ToString();     //ITEM_DAY
                        string valid_flg = FpSpread1.Sheets[0].Cells[r, 9].Value.ToString();    //valid_flg
                        string mro_item_cd1 = txt_mro_item_cd.Text.ToString();  

                      
                     

                        /*변경된값 가져오기*/

                        if (i == 0)
                            cg_mro_item_cd = e.EditValues[0].ToString();   
                        else
                            cg_mro_item_cd = mro_item_cd;

                        if (i == 1)
                            cg_item_group1 = e.EditValues[1].ToString();
                        else
                            cg_item_group1 = item_group1;

                        if (i == 2)
                            cg_item_group2 = e.EditValues[2].ToString();
                        else
                            cg_item_group2 = item_group2;

                        if (i == 3)
                            cg_item_nm = e.EditValues[3].ToString();
                        else
                            cg_item_nm = item_nm;

                        if (i == 4)
                            cg_spec = e.EditValues[4].ToString();
                        else
                            cg_spec = spec;


                        if (i == 5)
                            cg_unit = e.EditValues[5].ToString();
                        else
                            cg_unit = unit;
                        
                        if (i == 6)
                            cg_amt = e.EditValues[6].ToString();
                        else
                            cg_amt = amt;

                        if (i == 7)
                            cg_moq = e.EditValues[7].ToString();
                        else
                            cg_moq = moq;

                        if (i == 8)
                            cg_item_day = e.EditValues[8].ToString();
                        else
                            cg_item_day = item_day;

                        if (i == 9)
                            cg_valid_flg = e.EditValues[9].ToString();
                        else
                            cg_valid_flg = valid_flg;

                        sql = "update b_item_mro ";
                        sql = sql + "set ITEM_NM = '" + cg_item_nm + "',ITEM_GROUP1 = '" + cg_item_group1 + "',ITEM_GROUP2 = '" + cg_item_group2 + "', SPEC = '" + spec + "',";
                        sql = sql + "unit = '" + cg_unit + "',AMT = '" + cg_amt + "',MOQ = '" + cg_moq + "',ITEM_DAY = '" + cg_item_day + "',MRO_ITEM_CD = '" + cg_mro_item_cd + "',VALID_FLG = '" + cg_valid_flg + "',";
                        sql = sql + "updt_user_id ='" + Session["User"].ToString() + "',updt_dt =  getdate()";
                        sql = sql + " where mro_item_cd = '" + mro_item_cd1 + "' and  item_cd = '" + erp_item_cd + "'  ";
                        Execute_ERP(conn_erp,sql, "");
                    }
                }
            }

            /********************************* 조회 ************************************/
            protected void btn_search_Click(object sender, EventArgs e) //조회 버튼
            {
                string sql;

                string erp_item_cd = ddl_item_cd2.SelectedValue.ToString();             //ERP품목코드
                string mro_item_cd = txt_mro_item_cd.Text.ToString();                   //KEP품목코드


                sql = "select mro_item_cd, ITEM_GROUP1,ITEM_GROUP2, ITEM_NM, SPEC, UNIT, AMT, MOQ ,ITEM_DAY,VALID_FLG "+
                      " from b_item_mro " +
                      " where item_cd like '" + erp_item_cd + "' " +
                      "   and mro_item_cd like '" + mro_item_cd + "' ";

                sqlAdapter1 = new SqlDataAdapter(sql, conn_erp);

                sqlAdapter1.Fill(ds, "ds");

                FpSpread1.DataSource = ds;
                FpSpread1.DataBind();

            }

            /********************************* 삭제 ************************************/

            protected void btn_del_Click(object sender, EventArgs e)
            {
                System.Collections.IEnumerator enu = FpSpread1.ActiveSheetView.SelectionModel.GetEnumerator();
                FarPoint.Web.Spread.Model.CellRange cr;

                while (enu.MoveNext())
                {
                    cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                    int a = FpSpread1.Sheets[0].ActiveRow;
                    //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                    for (int i = 0; i < cr.RowCount; i++)
                    {
                        string mro_item_cd = FpSpread1.Sheets[0].Cells[FpSpread1.Sheets[0].ActiveRow, 0].Text;
                        string erp_item_cd = ddl_item_cd2.SelectedValue.ToString();
                        

                        string sql = "delete b_item_mro ";
                        sql = sql + " where mro_item_cd  ='" + mro_item_cd + "' and item_cd = '" + erp_item_cd + "' ";

                        if (Execute_ERP(conn_erp,sql, "") > 0)
                            FpSpread1.Sheets[0].Rows.Remove(FpSpread1.Sheets[0].ActiveRow, 1);
                       
                    }
                }

                MessageBox.ShowMessage("삭제되었습니다.", this.Page);
                btn_search_Click(null, null);
            }

           
    }
}