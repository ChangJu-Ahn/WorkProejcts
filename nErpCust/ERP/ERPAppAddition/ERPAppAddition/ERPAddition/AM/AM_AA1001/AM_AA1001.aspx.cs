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
using System.Drawing;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.QueryExe;


namespace ERPAppAddition.ERPAddition.AM.AM_AA1001 //일일운용자금실적(NEPES)
{
    public partial class AM_AA1001 : System.Web.UI.Page
    {

        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        DataSet ds = new DataSet();

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

                //MessageBox.ShowMessage(userid, this.Page);

                Session["User"] = userid;
                // rbl_view_type_SelectedIndexChanged(null, null);

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

            if (rbl_view_type.SelectedValue == "A") //일일실적등록 선택
            {
                Panel_regist_excel_grid.Visible = true;
                Panel_amt.Visible = false;
                Pane_excel.Visible = true;
                Panel_insert.Visible = false;
            }
            if (rbl_view_type.SelectedValue == "B") //경상/계열사 수입 등록 선택
            {
                Panel_regist_excel_grid.Visible = false;
                Panel_amt.Visible = true;
                Pane_excel.Visible = false;
                Panel_insert.Visible = true;
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e) //엑셀 업로드
        {
            if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
            {
                MessageBox.ShowMessage("'년도'를 입력하세요.", this.Page);

                return;
            }

            if (txt_mm == null || txt_mm.Text.Equals(""))
            {
                MessageBox.ShowMessage("'월'을 입력하세요.", this.Page);

                return;
            }

            if (txt_dd == null || txt_dd.Text.Equals(""))
            {
                MessageBox.ShowMessage("'일'을 입력하세요.", this.Page);

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

        protected void btn_save_Click(object sender, EventArgs e) //저장 버튼 클릭
        {
            try
            {

                //for (int i = 0; i < grid_regist_excel.Rows.Count; i++)
                for (int i = 0; i < 1; i++)
                {
                    chk_save_yn = 0;

                    string yyyy = txt_yyyy.Text.Trim();//년도
                    string mm = txt_mm.Text.Trim();//월
                    string dd = txt_dd.Text.Trim();//일
                    string LENDER_COLLET_1 = grid_regist_excel.Rows[0].Cells[i + 2].Text.Trim();
                    string LENDER_COLLET_2 = grid_regist_excel.Rows[1].Cells[i + 2].Text.Trim();
                    string LENDER_COLLET_3 = grid_regist_excel.Rows[2].Cells[i + 2].Text.Trim();
                    string LENDER_COLLET_4 = grid_regist_excel.Rows[3].Cells[i + 2].Text.Trim();
                    string LENDER_COLLET_5 = grid_regist_excel.Rows[4].Cells[i + 2].Text.Trim();
                    string LENDER_COLLET = grid_regist_excel.Rows[5].Cells[i + 2].Text.Trim();//매출대전회수
                    string SURTAX_REFUND_1 = grid_regist_excel.Rows[6].Cells[i + 2].Text.Trim();
                    string SURTAX_REFUND_2 = grid_regist_excel.Rows[7].Cells[i + 2].Text.Trim();
                    string SURTAX_REFUND_3 = grid_regist_excel.Rows[8].Cells[i + 2].Text.Trim();
                    string SURTAX_REFUND_4 = grid_regist_excel.Rows[9].Cells[i + 2].Text.Trim();
                    string SURTAX_REFUND_5 = grid_regist_excel.Rows[10].Cells[i + 2].Text.Trim();
                    string SURTAX_REFUND = grid_regist_excel.Rows[11].Cells[i + 2].Text.Trim();//부가세환급
                    string TARIFF_REFUND_1 = grid_regist_excel.Rows[12].Cells[i + 2].Text.Trim();
                    string TARIFF_REFUND_2 = grid_regist_excel.Rows[13].Cells[i + 2].Text.Trim();
                    string TARIFF_REFUND_3 = grid_regist_excel.Rows[14].Cells[i + 2].Text.Trim();
                    string TARIFF_REFUND_4 = grid_regist_excel.Rows[15].Cells[i + 2].Text.Trim();
                    string TARIFF_REFUND_5 = grid_regist_excel.Rows[16].Cells[i + 2].Text.Trim();
                    string TARIFF_REFUND = grid_regist_excel.Rows[17].Cells[i + 2].Text.Trim();//관세환급
                    string LEASE_INCOME_1 = grid_regist_excel.Rows[18].Cells[i + 2].Text.Trim();
                    string LEASE_INCOME_2 = grid_regist_excel.Rows[19].Cells[i + 2].Text.Trim();
                    string LEASE_INCOME_3 = grid_regist_excel.Rows[20].Cells[i + 2].Text.Trim();
                    string LEASE_INCOME_4 = grid_regist_excel.Rows[21].Cells[i + 2].Text.Trim();
                    string LEASE_INCOME_5 = grid_regist_excel.Rows[22].Cells[i + 2].Text.Trim();
                    string LEASE_INCOME = grid_regist_excel.Rows[23].Cells[i + 2].Text.Trim();// 임대보증금/임대료 입금
                    string IMPORT_INCOME_1 = grid_regist_excel.Rows[24].Cells[i + 2].Text.Trim();
                    string IMPORT_INCOME_2 = grid_regist_excel.Rows[25].Cells[i + 2].Text.Trim();
                    string IMPORT_INCOME_3 = grid_regist_excel.Rows[26].Cells[i + 2].Text.Trim();
                    string IMPORT_INCOME_4 = grid_regist_excel.Rows[27].Cells[i + 2].Text.Trim();
                    string IMPORT_INCOME_5 = grid_regist_excel.Rows[28].Cells[i + 2].Text.Trim();
                    string IMPORT_INCOME = grid_regist_excel.Rows[29].Cells[i + 2].Text.Trim();// 임대보증금/임대료 입금
                    string ETC_1 = grid_regist_excel.Rows[30].Cells[i + 2].Text.Trim();
                    string ETC_2 = grid_regist_excel.Rows[31].Cells[i + 2].Text.Trim();
                    string ETC_3 = grid_regist_excel.Rows[32].Cells[i + 2].Text.Trim();
                    string ETC_4 = grid_regist_excel.Rows[33].Cells[i + 2].Text.Trim();
                    string ETC_5 = grid_regist_excel.Rows[34].Cells[i + 2].Text.Trim();
                    string ETC = grid_regist_excel.Rows[35].Cells[i + 2].Text.Trim();// 기타
                    string BUSINESS_INCOME = grid_regist_excel.Rows[36].Cells[i + 2].Text.Trim();// [기타] 영업활동상의 자급수입
                    string IV_1 = grid_regist_excel.Rows[37].Cells[i + 2].Text.Trim();
                    string IV_2 = grid_regist_excel.Rows[38].Cells[i + 2].Text.Trim();
                    string IV_3 = grid_regist_excel.Rows[39].Cells[i + 2].Text.Trim();
                    string IV_4 = grid_regist_excel.Rows[40].Cells[i + 2].Text.Trim();
                    string IV_5 = grid_regist_excel.Rows[41].Cells[i + 2].Text.Trim();
                    string IV = grid_regist_excel.Rows[42].Cells[i + 2].Text.Trim();//원자재 매입대금 지급
                    string M_PAY_1 = grid_regist_excel.Rows[43].Cells[i + 2].Text.Trim();
                    string M_PAY_2 = grid_regist_excel.Rows[44].Cells[i + 2].Text.Trim();
                    string M_PAY_3 = grid_regist_excel.Rows[45].Cells[i + 2].Text.Trim();
                    string M_PAY_4 = grid_regist_excel.Rows[46].Cells[i + 2].Text.Trim();
                    string M_PAY_5 = grid_regist_excel.Rows[47].Cells[i + 2].Text.Trim();
                    string M_PAY = grid_regist_excel.Rows[48].Cells[i + 2].Text.Trim(); //급여와상여
                    string M_RETIRE_1 = grid_regist_excel.Rows[49].Cells[i + 2].Text.Trim();
                    string M_RETIRE_2 = grid_regist_excel.Rows[50].Cells[i + 2].Text.Trim();
                    string M_RETIRE_3 = grid_regist_excel.Rows[51].Cells[i + 2].Text.Trim();
                    string M_RETIRE_4 = grid_regist_excel.Rows[52].Cells[i + 2].Text.Trim();
                    string M_RETIRE_5 = grid_regist_excel.Rows[53].Cells[i + 2].Text.Trim();
                    string M_RETIRE = grid_regist_excel.Rows[54].Cells[i + 2].Text.Trim(); //퇴직급의 지급
                    string M_FOUNTAIN_1 = grid_regist_excel.Rows[55].Cells[i + 2].Text.Trim();
                    string M_FOUNTAIN_2 = grid_regist_excel.Rows[56].Cells[i + 2].Text.Trim();
                    string M_FOUNTAIN_3 = grid_regist_excel.Rows[57].Cells[i + 2].Text.Trim();
                    string M_FOUNTAIN_4 = grid_regist_excel.Rows[58].Cells[i + 2].Text.Trim();
                    string M_FOUNTAIN_5 = grid_regist_excel.Rows[59].Cells[i + 2].Text.Trim();
                    string M_FOUNTAIN = grid_regist_excel.Rows[60].Cells[i + 2].Text.Trim(); // 원천제세 납부
                    string M_WELRARE_1 = grid_regist_excel.Rows[61].Cells[i + 2].Text.Trim();
                    string M_WELRARE_2 = grid_regist_excel.Rows[62].Cells[i + 2].Text.Trim();
                    string M_WELRARE_3 = grid_regist_excel.Rows[63].Cells[i + 2].Text.Trim();
                    string M_WELRARE_4 = grid_regist_excel.Rows[64].Cells[i + 2].Text.Trim();
                    string M_WELRARE_5 = grid_regist_excel.Rows[65].Cells[i + 2].Text.Trim();
                    string M_WELRARE = grid_regist_excel.Rows[66].Cells[i + 2].Text.Trim(); //  법정복리비 납부
                    string PAYROLL_COSTS_EXPENSE = grid_regist_excel.Rows[67].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string PAYROLL_COSTS_EXPENSE_1 = grid_regist_excel.Rows[68].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string PAYROLL_COSTS_EXPENSE_2 = grid_regist_excel.Rows[69].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string PAYROLL_COSTS_EXPENSE_3 = grid_regist_excel.Rows[70].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string PAYROLL_COSTS_EXPENSE_4 = grid_regist_excel.Rows[71].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string PAYROLL_COSTS_EXPENSE_5 = grid_regist_excel.Rows[72].Cells[i + 2].Text.Trim(); // 인건비의 지급
                    string EXPENSE = grid_regist_excel.Rows[73].Cells[i + 2].Text.Trim();
                    string EXPENSE_1 = grid_regist_excel.Rows[74].Cells[i + 2].Text.Trim();
                    string EXPENSE_2 = grid_regist_excel.Rows[75].Cells[i + 2].Text.Trim();
                    string EXPENSE_3 = grid_regist_excel.Rows[76].Cells[i + 2].Text.Trim();
                    string EXPENSE_4 = grid_regist_excel.Rows[77].Cells[i + 2].Text.Trim();
                    string EXPENSE_5 = grid_regist_excel.Rows[78].Cells[i + 2].Text.Trim(); //사업부경비
                    string LEASE_EXPENSE = grid_regist_excel.Rows[79].Cells[i + 2].Text.Trim();
                    string LEASE_EXPENSE_1 = grid_regist_excel.Rows[80].Cells[i + 2].Text.Trim();
                    string LEASE_EXPENSE_2 = grid_regist_excel.Rows[81].Cells[i + 2].Text.Trim();
                    string LEASE_EXPENSE_3 = grid_regist_excel.Rows[82].Cells[i + 2].Text.Trim();
                    string LEASE_EXPENSE_4 = grid_regist_excel.Rows[83].Cells[i + 2].Text.Trim();
                    string LEASE_EXPENSE_5 = grid_regist_excel.Rows[84].Cells[i + 2].Text.Trim(); //임대보증금/임대료지급
                    string INSEREST = grid_regist_excel.Rows[85].Cells[i + 2].Text.Trim();
                    string INSEREST_1 = grid_regist_excel.Rows[86].Cells[i + 2].Text.Trim();
                    string INSEREST_2 = grid_regist_excel.Rows[87].Cells[i + 2].Text.Trim();
                    string INSEREST_3 = grid_regist_excel.Rows[88].Cells[i + 2].Text.Trim();
                    string INSEREST_4 = grid_regist_excel.Rows[89].Cells[i + 2].Text.Trim();
                    string INSEREST_5 = grid_regist_excel.Rows[90].Cells[i + 2].Text.Trim(); //지급이자
                    string BUSINESS_ETC = grid_regist_excel.Rows[91].Cells[i + 2].Text.Trim(); // [영업활동상의 자금수입] 기타
                    string BUSINESS_EXPENSE = grid_regist_excel.Rows[92].Cells[i + 2].Text.Trim(); // [영업활동상의 자금지출]
                    string BUSINESS_INFLOW = grid_regist_excel.Rows[93].Cells[i + 2].Text.Trim();  //영업상의 순 자금유입
                    string FIXED_ASSET_LAND = grid_regist_excel.Rows[94].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_LAND_1 = grid_regist_excel.Rows[95].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_LAND_2 = grid_regist_excel.Rows[96].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_LAND_3 = grid_regist_excel.Rows[97].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_LAND_4 = grid_regist_excel.Rows[98].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_LAND_5 = grid_regist_excel.Rows[99].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE = grid_regist_excel.Rows[100].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE_1 = grid_regist_excel.Rows[101].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE_2 = grid_regist_excel.Rows[102].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE_3 = grid_regist_excel.Rows[103].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE_4 = grid_regist_excel.Rows[104].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_MACHINE_5 = grid_regist_excel.Rows[105].Cells[i + 2].Text.Trim();//[고정자산의 매입] 기계장치 취득
                    string FIXED_ASSET_CAR = grid_regist_excel.Rows[106].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_CAR_1 = grid_regist_excel.Rows[107].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_CAR_2 = grid_regist_excel.Rows[108].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_CAR_3 = grid_regist_excel.Rows[109].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_CAR_4 = grid_regist_excel.Rows[110].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_CAR_5 = grid_regist_excel.Rows[111].Cells[i + 2].Text.Trim();//[고정자산의 매입] 차량 및 공기구 취득
                    string FIXED_ASSET_INCOME = grid_regist_excel.Rows[112].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_INCOME_1 = grid_regist_excel.Rows[113].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_INCOME_2 = grid_regist_excel.Rows[114].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_INCOME_3 = grid_regist_excel.Rows[115].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_INCOME_4 = grid_regist_excel.Rows[116].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_INCOME_5 = grid_regist_excel.Rows[117].Cells[i + 2].Text.Trim();//[고정자산의 매입] 보증금등의 회수
                    string FIXED_ASSET_ETC = grid_regist_excel.Rows[118].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_ETC1 = grid_regist_excel.Rows[119].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_ETC2 = grid_regist_excel.Rows[120].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_ETC3 = grid_regist_excel.Rows[121].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_ETC4 = grid_regist_excel.Rows[122].Cells[i + 2].Text.Trim();
                    string FIXED_ASSET_ETC5 = grid_regist_excel.Rows[123].Cells[i + 2].Text.Trim();//[고정자산의 매입] 기타
                    string FIXED_ASSET = grid_regist_excel.Rows[124].Cells[i + 2].Text.Trim();//[고정자산의 매입]
                    string LOAN_EXPENSE = grid_regist_excel.Rows[125].Cells[i + 2].Text.Trim();//대여금 지급
                    string FINANCIAL_OUT = grid_regist_excel.Rows[126].Cells[i + 2].Text.Trim();//장기금융상품의 매입
                    string INVEST_CASH = grid_regist_excel.Rows[127].Cells[i + 2].Text.Trim();//투자유가증권 / 출자금 지급
                    string BEWBORROW_0 = grid_regist_excel.Rows[128].Cells[i + 2].Text.Trim();
                    string BEWBORROW_1 = grid_regist_excel.Rows[129].Cells[i + 2].Text.Trim();
                    string BEWBORROW_2 = grid_regist_excel.Rows[130].Cells[i + 2].Text.Trim();
                    string BEWBORROW_3 = grid_regist_excel.Rows[131].Cells[i + 2].Text.Trim();
                    string BEWBORROW_4 = grid_regist_excel.Rows[132].Cells[i + 2].Text.Trim();
                    string BEWBORROW_5 = grid_regist_excel.Rows[133].Cells[i + 2].Text.Trim();// [투자활동의 자금흐름]신규차입
                    string BEWBORROW = grid_regist_excel.Rows[134].Cells[i + 2].Text.Trim();// [투자활동의 자금흐름]신규차입
                    string INVEST_USANCE_1 = grid_regist_excel.Rows[135].Cells[i + 2].Text.Trim();
                    string INVEST_USANCE_2 = grid_regist_excel.Rows[136].Cells[i + 2].Text.Trim();
                    string INVEST_USANCE_3 = grid_regist_excel.Rows[137].Cells[i + 2].Text.Trim();
                    string INVEST_USANCE_4 = grid_regist_excel.Rows[138].Cells[i + 2].Text.Trim();
                    string INVEST_USANCE_5 = grid_regist_excel.Rows[139].Cells[i + 2].Text.Trim();
                    string INVEST_USANCE = grid_regist_excel.Rows[140].Cells[i + 2].Text.Trim();// [투자활동의 자금흐름]USANCE 차입
                    string INCREASE_CASHIN = grid_regist_excel.Rows[141].Cells[i + 2].Text.Trim();// [투자활동의 자금흐름]증자 등
                    string LOAN_CD_0 = grid_regist_excel.Rows[142].Cells[i + 2].Text.Trim();
                    string LOAN_CD1 = grid_regist_excel.Rows[143].Cells[i + 2].Text.Trim();
                    string LOAN_CD2 = grid_regist_excel.Rows[144].Cells[i + 2].Text.Trim();
                    string LOAN_CD3 = grid_regist_excel.Rows[145].Cells[i + 2].Text.Trim();
                    string LOAN_CD4 = grid_regist_excel.Rows[146].Cells[i + 2].Text.Trim();
                    string LOAN_CD5 = grid_regist_excel.Rows[147].Cells[i + 2].Text.Trim();
                    string LOAN_CD = grid_regist_excel.Rows[148].Cells[i + 2].Text.Trim(); //차입금의 상환
                    string USANCE_CD1 = grid_regist_excel.Rows[149].Cells[i + 2].Text.Trim();
                    string USANCE_CD2 = grid_regist_excel.Rows[150].Cells[i + 2].Text.Trim();
                    string USANCE_CD3 = grid_regist_excel.Rows[151].Cells[i + 2].Text.Trim();
                    string USANCE_CD4 = grid_regist_excel.Rows[152].Cells[i + 2].Text.Trim();
                    string USANCE_CD5 = grid_regist_excel.Rows[153].Cells[i + 2].Text.Trim();
                    string USANCE_CD = grid_regist_excel.Rows[154].Cells[i + 2].Text.Trim();
                    string DIVIDEND_CASHOUT = grid_regist_excel.Rows[155].Cells[i + 2].Text.Trim(); //배당금/자기주식 취득 외
                    string CASH_OUT = grid_regist_excel.Rows[156].Cells[i + 2].Text.Trim(); //재무적지출
                    string CASH_IN = grid_regist_excel.Rows[157].Cells[i + 2].Text.Trim(); //재무적 순 자금 유입
                    string LAST_AMT = grid_regist_excel.Rows[158].Cells[i + 2].Text.Trim(); //전월이월 자금
                    string LAST_LESS_AMT = grid_regist_excel.Rows[159].Cells[i + 2].Text.Trim(); //자금의 과부족
                    string AFTER_AMT = grid_regist_excel.Rows[160].Cells[i + 2].Text.Trim(); // 차월이월 자금


                    string sql = "select count(dd) from A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text.Trim() + "' AND MM = '" + txt_mm.Text.Trim() + "' ";
                    sql += "and DD = '" + txt_dd.Text.Trim() + "' ";


                    if (Execute_ERP(conn_erp, sql, "check") > 0)//기존자료가 있으면 삭제
                    {

                        sql = "delete A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text.Trim() + "' AND MM = '" + txt_mm.Text.Trim() + "' ";
                        sql += "and DD = '" + txt_dd.Text.Trim() + "' ";


                        if (Execute_ERP(conn_erp, sql, "") > 0)
                        {
                            sql = "INSERT INTO A_DAILY_AMT VALUES('" + yyyy + "' ,'" + mm + "' ,'" + dd + "',";
                            sql += "'" + LENDER_COLLET_1 + "','" + LENDER_COLLET_2 + "' ,'" + LENDER_COLLET_3 + "','" + LENDER_COLLET_4 + "','" + LENDER_COLLET_5 + "','" + LENDER_COLLET + "',"; //매출대전회수
                            sql += "'" + SURTAX_REFUND_1 + "','" + SURTAX_REFUND_2 + "' ,'" + SURTAX_REFUND_3 + "','" + SURTAX_REFUND_4 + "','" + SURTAX_REFUND_5 + "','" + SURTAX_REFUND + "',"; //부가세환급
                            sql += "'" + TARIFF_REFUND_1 + "','" + TARIFF_REFUND_2 + "' ,'" + TARIFF_REFUND_3 + "','" + TARIFF_REFUND_4 + "','" + TARIFF_REFUND_5 + "','" + TARIFF_REFUND + "',"; //관세환급
                            sql += "'" + LEASE_INCOME_1 + "','" + LEASE_INCOME_2 + "' ,'" + LEASE_INCOME_3 + "','" + LEASE_INCOME_4 + "','" + LEASE_INCOME_5 + "','" + LEASE_INCOME + "',"; //임대보증금/임대료 입금
                            sql += "'" + IMPORT_INCOME_1 + "','" + IMPORT_INCOME_2 + "' ,'" + IMPORT_INCOME_3 + "','" + IMPORT_INCOME_4 + "','" + IMPORT_INCOME_5 + "','" + IMPORT_INCOME + "',"; //수입이자입금
                            sql += "'" + ETC_1 + "','" + ETC_2 + "' ,'" + ETC_3 + "','" + ETC_4 + "','" + ETC_5 + "','" + ETC + "',"; //기타
                            sql += "'" + BUSINESS_INCOME + "',"; //[기타] 영업활동상의 자급수입
                            sql += "'" + IV_1 + "','" + IV_2 + "' ,'" + IV_3 + "','" + IV_4 + "','" + IV_5 + "','" + IV + "',"; //원자재 매입대금 지급
                            sql += "'" + M_PAY_1 + "','" + M_PAY_2 + "' ,'" + M_PAY_3 + "','" + M_PAY_4 + "','" + M_PAY_5 + "','" + M_PAY + "',"; //급여와 상여
                            sql += "'" + M_RETIRE_1 + "','" + M_RETIRE_2 + "' ,'" + M_RETIRE_3 + "','" + M_RETIRE_4 + "','" + M_RETIRE_5 + "','" + M_RETIRE + "',"; //퇴직급의 지급
                            sql += "'" + M_FOUNTAIN_1 + "','" + M_FOUNTAIN_2 + "' ,'" + M_FOUNTAIN_3 + "','" + M_FOUNTAIN_4 + "','" + M_FOUNTAIN_5 + "','" + M_FOUNTAIN + "',"; //원천제세 납부
                            sql += "'" + M_WELRARE_1 + "','" + M_WELRARE_2 + "' ,'" + M_WELRARE_3 + "','" + M_WELRARE_4 + "','" + M_WELRARE_5 + "','" + M_WELRARE + "',"; //법정복리비 납부
                            sql += "'" + PAYROLL_COSTS_EXPENSE + "',"; //인건비의 지급
                            sql += "'" + PAYROLL_COSTS_EXPENSE_1 + "','" + PAYROLL_COSTS_EXPENSE_2 + "' ,'" + PAYROLL_COSTS_EXPENSE_3 + "','" + PAYROLL_COSTS_EXPENSE_4 + "','" + PAYROLL_COSTS_EXPENSE_5 + "',"; //인건비의 지급
                            sql += "'" + EXPENSE + "','" + EXPENSE_1 + "' ,'" + EXPENSE_2 + "','" + EXPENSE_3 + "','" + EXPENSE_4 + "','" + EXPENSE_5 + "',"; //사업부경비
                            sql += "'" + LEASE_EXPENSE + "','" + LEASE_EXPENSE_1 + "' ,'" + LEASE_EXPENSE_2 + "','" + LEASE_EXPENSE_3 + "','" + LEASE_EXPENSE_4 + "','" + LEASE_EXPENSE_5 + "',"; //임대보증금/임대료지급
                            sql += "'" + INSEREST + "','" + INSEREST_1 + "' ,'" + INSEREST_2 + "','" + INSEREST_3 + "','" + INSEREST_4 + "','" + INSEREST_5 + "',"; //지급이자
                            sql += "'" + BUSINESS_ETC + "','" + BUSINESS_EXPENSE + "' ,'" + BUSINESS_INFLOW + "',"; //영업활동상의 자금수입] 기타,[영업활동상의 자금지출]
                            sql += "'" + FIXED_ASSET_LAND + "','" + FIXED_ASSET_LAND_1 + "' ,'" + FIXED_ASSET_LAND_2 + "','" + FIXED_ASSET_LAND_3 + "' ,'" + FIXED_ASSET_LAND_4 + "','" + FIXED_ASSET_LAND_5 + "', ";
                            sql += "'" + FIXED_ASSET_MACHINE + "','" + FIXED_ASSET_MACHINE_1 + "' ,'" + FIXED_ASSET_MACHINE_2 + "','" + FIXED_ASSET_MACHINE_3 + "','" + FIXED_ASSET_MACHINE_4 + "','" + FIXED_ASSET_MACHINE_5 + "',"; //[고정자산의 매입] 기계장치 취득
                            sql += "'" + FIXED_ASSET_CAR + "','" + FIXED_ASSET_CAR_1 + "' ,'" + FIXED_ASSET_CAR_2 + "','" + FIXED_ASSET_CAR_3 + "' ,'" + FIXED_ASSET_CAR_4 + "','" + FIXED_ASSET_CAR_5 + "',"; //[고정자산의 매입] 차량 및 공기구 취득
                            sql += "'" + FIXED_ASSET_INCOME + "','" + FIXED_ASSET_INCOME_1 + "' ,'" + FIXED_ASSET_INCOME_2 + "','" + FIXED_ASSET_INCOME_3 + "' ,'" + FIXED_ASSET_INCOME_4 + "','" + FIXED_ASSET_INCOME_5 + "',"; //[고정자산의 매입] 보증금등의 회수
                            sql += "'" + FIXED_ASSET_ETC + "','" + FIXED_ASSET_ETC1 + "' ,'" + FIXED_ASSET_ETC2 + "','" + FIXED_ASSET_ETC3 + "' ,'" + FIXED_ASSET_ETC4 + "','" + FIXED_ASSET_ETC5 + "', "; //[고정자산의 매입] 기타
                            sql += "'" + FIXED_ASSET + "','" + LOAN_EXPENSE + "' ,'" + FINANCIAL_OUT + "','" + INVEST_CASH + "',"; //[고정자산의 매입]대여금,장기금융상품의 매입,투자유가증권 / 출자금 지급
                            sql += "'" + BEWBORROW_0 + "','" + BEWBORROW_1 + "' ,'" + BEWBORROW_2 + "','" + BEWBORROW_3 + "' ,'" + BEWBORROW_4 + "','" + BEWBORROW_5 + "','" + BEWBORROW + "',"; // [투자활동의 자금흐름]신규차입
                            sql += "'" + INVEST_USANCE_1 + "','" + INVEST_USANCE_2 + "' ,'" + INVEST_USANCE_3 + "','" + INVEST_USANCE_4 + "' ,'" + INVEST_USANCE_5 + "','" + INVEST_USANCE + "',"; //[투자활동의 자금흐름]USANCE 차입
                            sql += "'" + INCREASE_CASHIN + "',";
                            sql += "'" + LOAN_CD_0 + "','" + LOAN_CD1 + "' ,'" + LOAN_CD2 + "','" + LOAN_CD3 + "' ,'" + LOAN_CD4 + "','" + LOAN_CD5 + "','" + LOAN_CD + "', "; //차입금의 상환
                            sql += "'" + USANCE_CD1 + "','" + USANCE_CD2 + "' ,'" + USANCE_CD3 + "','" + USANCE_CD4 + "' ,'" + USANCE_CD5 + "','" + USANCE_CD + "',"; //
                            sql += "'" + DIVIDEND_CASHOUT + "',"; //배당금/자기주식 취득 외
                            sql += "'" + CASH_OUT + "','" + CASH_IN + "' ,'" + LAST_AMT + "','" + LAST_LESS_AMT + "' ,'" + AFTER_AMT + "' ,'',"; //재무적지출~비고
                            sql += "'" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Session["User"].ToString() + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' )";


                            if (Execute_ERP(conn_erp, sql, "") > 0)
                                chk_save_yn += 1;
                        }
                    }
                    else
                    {
                        //엑셀데이타 값이 없는것은 제외:
                        if ((yyyy == "" || yyyy == null || yyyy == "&nbsp;") || (dd == "" || dd == null || dd == "&nbsp;") || (mm == "" || mm == null || mm == "&nbsp;")
                            || (LENDER_COLLET_1 == "" || LENDER_COLLET_1 == null || LENDER_COLLET_1 == "&nbsp;") || (LENDER_COLLET_2 == "" || LENDER_COLLET_2 == null || LENDER_COLLET_2 == "&nbsp;") || (LENDER_COLLET_3 == "" || LENDER_COLLET_3 == null || LENDER_COLLET_3 == "&nbsp;")
                            || (LENDER_COLLET_4 == "" || LENDER_COLLET_4 == null || LENDER_COLLET_4 == "&nbsp;") || (LENDER_COLLET_5 == "" || LENDER_COLLET_5 == null || LENDER_COLLET_5 == "&nbsp;") || (LENDER_COLLET == "" || LENDER_COLLET == null || LENDER_COLLET == "&nbsp;") //매출대전회수
                            || (SURTAX_REFUND_1 == "" || SURTAX_REFUND_1 == null || SURTAX_REFUND_1 == "&nbsp;") || (SURTAX_REFUND_2 == "" || SURTAX_REFUND_2 == null || SURTAX_REFUND_2 == "&nbsp;") || (SURTAX_REFUND_3 == "" || SURTAX_REFUND_3 == null || SURTAX_REFUND_3 == "&nbsp;")
                            || (SURTAX_REFUND_4 == "" || SURTAX_REFUND_4 == null || SURTAX_REFUND_4 == "&nbsp;") || (SURTAX_REFUND_5 == "" || SURTAX_REFUND_5 == null || SURTAX_REFUND_5 == "&nbsp;") || (SURTAX_REFUND == "" || SURTAX_REFUND == null || SURTAX_REFUND == "&nbsp;") //부가세환급
                            || (TARIFF_REFUND_1 == "" || TARIFF_REFUND_1 == null || TARIFF_REFUND_1 == "&nbsp;") || (TARIFF_REFUND_2 == "" || TARIFF_REFUND_2 == null || TARIFF_REFUND_2 == "&nbsp;") || (SURTAX_REFUND_3 == "" || SURTAX_REFUND_3 == null || SURTAX_REFUND_3 == "&nbsp;")
                            || (TARIFF_REFUND_4 == "" || TARIFF_REFUND_4 == null || TARIFF_REFUND_4 == "&nbsp;") || (TARIFF_REFUND_5 == "" || TARIFF_REFUND_5 == null || TARIFF_REFUND_5 == "&nbsp;") || (TARIFF_REFUND == "" || TARIFF_REFUND == null || TARIFF_REFUND == "&nbsp;") //임대보증금/임대료 입금
                            || (LEASE_INCOME_1 == "" || LEASE_INCOME_1 == null || LEASE_INCOME_1 == "&nbsp;") || (LEASE_INCOME_2 == "" || LEASE_INCOME_2 == null || LEASE_INCOME_2 == "&nbsp;") || (LEASE_INCOME_3 == "" || LEASE_INCOME_3 == null || LEASE_INCOME_3 == "&nbsp;")
                            || (LEASE_INCOME_4 == "" || LEASE_INCOME_4 == null || LEASE_INCOME_4 == "&nbsp;") || (LEASE_INCOME_5 == "" || LEASE_INCOME_5 == null || LEASE_INCOME_5 == "&nbsp;") || (LEASE_INCOME == "" || LEASE_INCOME == null || LEASE_INCOME == "&nbsp;") //임대보증금/임대료 입금
                            || (ETC_1 == "" || ETC_1 == null || ETC_1 == "&nbsp;") || (ETC_2 == "" || ETC_2 == null || ETC_2 == "&nbsp;") || (ETC_3 == "" || ETC_3 == null || ETC_3 == "&nbsp;")
                            || (ETC_4 == "" || ETC_4 == null || ETC_4 == "&nbsp;") || (ETC_5 == "" || ETC_5 == null || ETC_5 == "&nbsp;") || (ETC == "" || ETC == null || ETC == "&nbsp;") //기타
                            || (BUSINESS_INCOME == "" || BUSINESS_INCOME == null || BUSINESS_INCOME == "&nbsp;")  //[기타] 영업활동상의 자급수입
                            || (IV_1 == "" || IV_1 == null || IV_1 == "&nbsp;") || (IV_2 == "" || IV_2 == null || IV_2 == "&nbsp;") || (IV_3 == "" || IV_3 == null || IV_3 == "&nbsp;")
                            || (IV_4 == "" || IV_4 == null || IV_4 == "&nbsp;") || (IV_5 == "" || IV_5 == null || IV_5 == "&nbsp;") || (IV == "" || IV == null || IV == "&nbsp;") //원자재 매입대금 지급
                            || (M_PAY_1 == "" || M_PAY_1 == null || M_PAY_1 == "&nbsp;") || (M_PAY_2 == "" || M_PAY_2 == null || M_PAY_2 == "&nbsp;") || (M_PAY_3 == "" || M_PAY_3 == null || M_PAY_3 == "&nbsp;")
                            || (M_PAY_4 == "" || M_PAY_4 == null || M_PAY_4 == "&nbsp;") || (M_PAY_5 == "" || M_PAY_5 == null || M_PAY_5 == "&nbsp;") || (M_PAY == "" || M_PAY == null || M_PAY == "&nbsp;") //급여와 상여
                            || (M_RETIRE_1 == "" || M_RETIRE_1 == null || M_RETIRE_1 == "&nbsp;") || (M_RETIRE_2 == "" || M_RETIRE_2 == null || M_RETIRE_2 == "&nbsp;") || (M_RETIRE_3 == "" || M_RETIRE_3 == null || M_RETIRE_3 == "&nbsp;")
                            || (M_RETIRE_4 == "" || M_RETIRE_4 == null || M_RETIRE_4 == "&nbsp;") || (M_RETIRE_5 == "" || M_RETIRE_5 == null || M_RETIRE_5 == "&nbsp;") || (M_RETIRE == "" || M_RETIRE == null || M_RETIRE == "&nbsp;") //퇴직금의 직급
                            || (M_FOUNTAIN_1 == "" || M_FOUNTAIN_1 == null || M_FOUNTAIN_1 == "&nbsp;") || (M_FOUNTAIN_2 == "" || M_FOUNTAIN_2 == null || M_FOUNTAIN_2 == "&nbsp;") || (M_FOUNTAIN_3 == "" || M_FOUNTAIN_3 == null || M_FOUNTAIN_3 == "&nbsp;")
                            || (M_FOUNTAIN_4 == "" || M_FOUNTAIN_4 == null || M_FOUNTAIN_4 == "&nbsp;") || (M_FOUNTAIN_5 == "" || M_FOUNTAIN_5 == null || M_FOUNTAIN_5 == "&nbsp;") || (M_FOUNTAIN == "" || M_FOUNTAIN == null || M_FOUNTAIN == "&nbsp;") //원천제세 납부
                            || (M_WELRARE_1 == "" || M_WELRARE_1 == null || M_WELRARE_1 == "&nbsp;") || (M_WELRARE_2 == "" || M_WELRARE_2 == null || M_WELRARE_2 == "&nbsp;") || (M_WELRARE_3 == "" || M_WELRARE_3 == null || M_WELRARE_3 == "&nbsp;")
                            || (M_WELRARE_4 == "" || M_WELRARE_4 == null || M_WELRARE_4 == "&nbsp;") || (M_WELRARE_5 == "" || M_WELRARE_5 == null || M_WELRARE_5 == "&nbsp;") || (M_WELRARE == "" || M_WELRARE == null || M_WELRARE == "&nbsp;") //법정복리비 납부
                            || (PAYROLL_COSTS_EXPENSE == "" || PAYROLL_COSTS_EXPENSE == null || PAYROLL_COSTS_EXPENSE == "&nbsp;") //인건비의 지급
                            || (PAYROLL_COSTS_EXPENSE_1 == "" || PAYROLL_COSTS_EXPENSE_1 == null || PAYROLL_COSTS_EXPENSE_1 == "&nbsp;") //인건비의 지급1
                            || (PAYROLL_COSTS_EXPENSE_2 == "" || PAYROLL_COSTS_EXPENSE_2 == null || PAYROLL_COSTS_EXPENSE_2 == "&nbsp;") //인건비의 지급2
                            || (PAYROLL_COSTS_EXPENSE_3 == "" || PAYROLL_COSTS_EXPENSE_3 == null || PAYROLL_COSTS_EXPENSE_3 == "&nbsp;") //인건비의 지급3
                            || (PAYROLL_COSTS_EXPENSE_4 == "" || PAYROLL_COSTS_EXPENSE_4 == null || PAYROLL_COSTS_EXPENSE_4 == "&nbsp;") //인건비의 지급4
                            || (PAYROLL_COSTS_EXPENSE_5 == "" || PAYROLL_COSTS_EXPENSE_5 == null || PAYROLL_COSTS_EXPENSE_5 == "&nbsp;") //인건비의 지급5
                            || (EXPENSE == "" || EXPENSE == null || EXPENSE == "&nbsp;") || (EXPENSE_1 == "" || EXPENSE_1 == null || EXPENSE_1 == "&nbsp;") || (EXPENSE_2 == "" || EXPENSE_2 == null || EXPENSE_2 == "&nbsp;")
                            || (EXPENSE_3 == "" || EXPENSE_3 == null || EXPENSE_3 == "&nbsp;") || (EXPENSE_4 == "" || EXPENSE_4 == null || EXPENSE_4 == "&nbsp;") || (EXPENSE_5 == "" || EXPENSE_5 == null || EXPENSE_5 == "&nbsp;") //사업부경비
                            || (LEASE_EXPENSE == "" || LEASE_EXPENSE == null || LEASE_EXPENSE == "&nbsp;") || (LEASE_EXPENSE_1 == "" || LEASE_EXPENSE_1 == null || LEASE_EXPENSE_1 == "&nbsp;") || (LEASE_EXPENSE_2 == "" || LEASE_EXPENSE_2 == null || LEASE_EXPENSE_2 == "&nbsp;")
                            || (LEASE_EXPENSE_3 == "" || LEASE_EXPENSE_3 == null || LEASE_EXPENSE_3 == "&nbsp;") || (LEASE_EXPENSE_4 == "" || LEASE_EXPENSE_4 == null || LEASE_EXPENSE_4 == "&nbsp;") || (LEASE_EXPENSE_5 == "" || LEASE_EXPENSE_5 == null || LEASE_EXPENSE_5 == "&nbsp;")//임대보증금/임대료지급
                            || (INSEREST == "" || INSEREST == null || INSEREST == "&nbsp;") || (INSEREST_1 == "" || INSEREST_1 == null || INSEREST_1 == "&nbsp;") || (INSEREST_2 == "" || INSEREST_2 == null || INSEREST_2 == "&nbsp;")
                            || (INSEREST_3 == "" || INSEREST_3 == null || INSEREST_3 == "&nbsp;") || (INSEREST_4 == "" || INSEREST_4 == null || INSEREST_4 == "&nbsp;") || (INSEREST_5 == "" || INSEREST_5 == null || INSEREST_5 == "&nbsp;")//지급이자
                            || (INSEREST == "" || INSEREST == null || INSEREST == "&nbsp;") || (INSEREST_1 == "" || INSEREST_1 == null || INSEREST_1 == "&nbsp;") || (INSEREST_2 == "" || INSEREST_2 == null || INSEREST_2 == "&nbsp;")
                            || (BUSINESS_ETC == "" || BUSINESS_ETC == null || BUSINESS_ETC == "&nbsp;") || (BUSINESS_EXPENSE == "" || BUSINESS_EXPENSE == null || BUSINESS_EXPENSE == "&nbsp;") || (BUSINESS_INFLOW == "" || BUSINESS_INFLOW == null || BUSINESS_INFLOW == "&nbsp;")//영업활동상의 자금수입] 기타,[영업활동상의 자금지출]
                            || (FIXED_ASSET_LAND == "" || FIXED_ASSET_LAND == null || FIXED_ASSET_LAND == "&nbsp;") || (FIXED_ASSET_LAND_1 == "" || FIXED_ASSET_LAND_1 == null || FIXED_ASSET_LAND_1 == "&nbsp;") || (FIXED_ASSET_LAND_2 == "" || FIXED_ASSET_LAND_2 == null || FIXED_ASSET_LAND_2 == "&nbsp;")
                            || (FIXED_ASSET_LAND_3 == "" || FIXED_ASSET_LAND_3 == null || FIXED_ASSET_LAND_3 == "&nbsp;") || (FIXED_ASSET_LAND_4 == "" || FIXED_ASSET_LAND_4 == null || FIXED_ASSET_LAND_4 == "&nbsp;") || (FIXED_ASSET_LAND_5 == "" || FIXED_ASSET_LAND_5 == null || FIXED_ASSET_LAND_5 == "&nbsp;")
                            || (FIXED_ASSET_MACHINE == "" || FIXED_ASSET_MACHINE == null || FIXED_ASSET_MACHINE == "&nbsp;") || (FIXED_ASSET_LAND_1 == "" || FIXED_ASSET_LAND_1 == null || FIXED_ASSET_LAND_1 == "&nbsp;") || (FIXED_ASSET_LAND_2 == "" || FIXED_ASSET_LAND_2 == null || FIXED_ASSET_LAND_2 == "&nbsp;")
                            || (FIXED_ASSET_MACHINE_3 == "" || FIXED_ASSET_MACHINE_3 == null || FIXED_ASSET_MACHINE_3 == "&nbsp;") || (FIXED_ASSET_MACHINE_4 == "" || FIXED_ASSET_MACHINE_4 == null || FIXED_ASSET_MACHINE_4 == "&nbsp;") || (FIXED_ASSET_MACHINE_5 == "" || FIXED_ASSET_MACHINE_5 == null || FIXED_ASSET_MACHINE_5 == "&nbsp;")//[고정자산의 매입] 기계장치 취득
                            || (FIXED_ASSET_CAR == "" || FIXED_ASSET_CAR == null || FIXED_ASSET_CAR == "&nbsp;") || (FIXED_ASSET_CAR_1 == "" || FIXED_ASSET_CAR_1 == null || FIXED_ASSET_CAR_1 == "&nbsp;") || (FIXED_ASSET_CAR_2 == "" || FIXED_ASSET_CAR_2 == null || FIXED_ASSET_CAR_2 == "&nbsp;")
                            || (FIXED_ASSET_CAR_3 == "" || FIXED_ASSET_MACHINE_3 == null || FIXED_ASSET_CAR_3 == "&nbsp;") || (FIXED_ASSET_CAR_4 == "" || FIXED_ASSET_CAR_4 == null || FIXED_ASSET_CAR_4 == "&nbsp;") || (FIXED_ASSET_CAR_5 == "" || FIXED_ASSET_CAR_5 == null || FIXED_ASSET_CAR_5 == "&nbsp;")//[고정자산의 매입] 차량 및 공기구 취득
                            || (FIXED_ASSET_INCOME == "" || FIXED_ASSET_INCOME == null || FIXED_ASSET_INCOME == "&nbsp;") || (FIXED_ASSET_INCOME_1 == "" || FIXED_ASSET_INCOME_1 == null || FIXED_ASSET_INCOME_1 == "&nbsp;") || (FIXED_ASSET_INCOME_2 == "" || FIXED_ASSET_INCOME_2 == null || FIXED_ASSET_INCOME_2 == "&nbsp;")
                            || (FIXED_ASSET_INCOME_3 == "" || FIXED_ASSET_INCOME_3 == null || FIXED_ASSET_INCOME_3 == "&nbsp;") || (FIXED_ASSET_INCOME_4 == "" || FIXED_ASSET_INCOME_4 == null || FIXED_ASSET_INCOME_4 == "&nbsp;") || (FIXED_ASSET_INCOME_5 == "" || FIXED_ASSET_INCOME_5 == null || FIXED_ASSET_INCOME_5 == "&nbsp;")//[고정자산의 매입] 보증금등의 회수
                            || (FIXED_ASSET_ETC == "" || FIXED_ASSET_ETC == null || FIXED_ASSET_ETC == "&nbsp;") || (FIXED_ASSET_ETC1 == "" || FIXED_ASSET_ETC1 == null || FIXED_ASSET_ETC1 == "&nbsp;") || (FIXED_ASSET_ETC2 == "" || FIXED_ASSET_ETC2 == null || FIXED_ASSET_ETC2 == "&nbsp;")
                            || (FIXED_ASSET_ETC3 == "" || FIXED_ASSET_ETC3 == null || FIXED_ASSET_ETC3 == "&nbsp;") || (FIXED_ASSET_ETC4 == "" || FIXED_ASSET_ETC4 == null || FIXED_ASSET_ETC4 == "&nbsp;") || (FIXED_ASSET_ETC5 == "" || FIXED_ASSET_ETC5 == null || FIXED_ASSET_ETC5 == "&nbsp;")//[고정자산의 매입] 기타
                            || (FIXED_ASSET == "" || FIXED_ASSET == null || FIXED_ASSET == "&nbsp;") || (LOAN_EXPENSE == "" || LOAN_EXPENSE == null || LOAN_EXPENSE == "&nbsp;") || (FINANCIAL_OUT == "" || FINANCIAL_OUT == null || FINANCIAL_OUT == "&nbsp;")
                            || (INVEST_CASH == "" || INVEST_CASH == null || INVEST_CASH == "&nbsp;")//[고정자산의 매입],대여금,장기금융상품의 매입,투자유가증권 / 출자금 지급
                            || (BEWBORROW_0 == "" || BEWBORROW_0 == null || BEWBORROW_0 == "&nbsp;") || (BEWBORROW_1 == "" || BEWBORROW_1 == null || BEWBORROW_1 == "&nbsp;") || (BEWBORROW_2 == "" || BEWBORROW_2 == null || BEWBORROW_2 == "&nbsp;")
                            || (BEWBORROW_3 == "" || BEWBORROW_3 == null || BEWBORROW_3 == "&nbsp;") || (BEWBORROW_4 == "" || BEWBORROW_4 == null || BEWBORROW_4 == "&nbsp;") || (BEWBORROW_5 == "" || BEWBORROW_5 == null || BEWBORROW_5 == "&nbsp;")// [투자활동의 자금흐름]신규차입
                            || (BEWBORROW == "" || BEWBORROW == null || BEWBORROW == "&nbsp;")//[투자활동의 자금흐름]신규차입
                            || (INVEST_USANCE_1 == "" || INVEST_USANCE_1 == null || INVEST_USANCE_1 == "&nbsp;") || (INVEST_USANCE_2 == "" || INVEST_USANCE_2 == null || INVEST_USANCE_2 == "&nbsp;") || (INVEST_USANCE_3 == "" || INVEST_USANCE_3 == null || INVEST_USANCE_3 == "&nbsp;")
                            || (INVEST_USANCE_4 == "" || INVEST_USANCE_4 == null || INVEST_USANCE_4 == "&nbsp;") || (INVEST_USANCE_5 == "" || INVEST_USANCE_5 == null || INVEST_USANCE_5 == "&nbsp;") || (INVEST_USANCE == "" || INVEST_USANCE == null || INVEST_USANCE == "&nbsp;") //[투자활동의 자금흐름]USANCE 차입
                            || (INCREASE_CASHIN == "" || INCREASE_CASHIN == null || INCREASE_CASHIN == "&nbsp;")
                            || (LOAN_CD_0 == "" || LOAN_CD_0 == null || LOAN_CD_0 == "&nbsp;") || (LOAN_CD1 == "" || LOAN_CD1 == null || LOAN_CD1 == "&nbsp;") || (LOAN_CD2 == "" || LOAN_CD2 == null || LOAN_CD2 == "&nbsp;")
                            || (LOAN_CD3 == "" || LOAN_CD3 == null || LOAN_CD3 == "&nbsp;") || (LOAN_CD4 == "" || LOAN_CD4 == null || LOAN_CD4 == "&nbsp;") || (LOAN_CD5 == "" || LOAN_CD5 == null || LOAN_CD5 == "&nbsp;") || (LOAN_CD == "" || LOAN_CD == null || LOAN_CD == "&nbsp;") //차입금의 상환
                            || (USANCE_CD1 == "" || USANCE_CD1 == null || USANCE_CD1 == "&nbsp;") || (USANCE_CD2 == "" || USANCE_CD2 == null || USANCE_CD2 == "&nbsp;") || (USANCE_CD3 == "" || USANCE_CD3 == null || USANCE_CD3 == "&nbsp;")
                            || (USANCE_CD4 == "" || USANCE_CD4 == null || USANCE_CD4 == "&nbsp;") || (USANCE_CD5 == "" || USANCE_CD5 == null || USANCE_CD5 == "&nbsp;") || (USANCE_CD == "" || USANCE_CD == null || USANCE_CD == "&nbsp;") //[투자활동의 자금흐름]USANCE 차입
                            || (DIVIDEND_CASHOUT == "" || DIVIDEND_CASHOUT == null || DIVIDEND_CASHOUT == "&nbsp;")//배당금/자기주식 취득 외
                            || (CASH_OUT == "" || CASH_OUT == null || CASH_OUT == "&nbsp;") || (CASH_IN == "" || CASH_IN == null || CASH_IN == "&nbsp;") || (LAST_AMT == "" || LAST_AMT == null || LAST_AMT == "&nbsp;")
                            || (LAST_LESS_AMT == "" || LAST_LESS_AMT == null || LAST_LESS_AMT == "&nbsp;") || (AFTER_AMT == "" || AFTER_AMT == null || AFTER_AMT == "&nbsp;"))
                        {
                            MessageBox.ShowMessage(i.ToString() + "빈 값이 있는 칸이 있습니다. ", this);
                            return;

                        }

                        else
                        {
                            sql = "INSERT INTO A_DAILY_AMT VALUES('" + yyyy + "' ,'" + mm + "' ,'" + dd + "',";
                            //   sql = " values('" + yyyy + "' ,'" + mm + "' ,'" + dd + "',";
                            sql += "'" + LENDER_COLLET_1 + "','" + LENDER_COLLET_2 + "' ,'" + LENDER_COLLET_3 + "','" + LENDER_COLLET_4 + "','" + LENDER_COLLET_5 + "','" + LENDER_COLLET + "',"; //매출대전회수
                            sql += "'" + SURTAX_REFUND_1 + "','" + SURTAX_REFUND_2 + "' ,'" + SURTAX_REFUND_3 + "','" + SURTAX_REFUND_4 + "','" + SURTAX_REFUND_5 + "','" + SURTAX_REFUND + "',"; //부가세환급
                            sql += "'" + TARIFF_REFUND_1 + "','" + TARIFF_REFUND_2 + "' ,'" + TARIFF_REFUND_3 + "','" + TARIFF_REFUND_4 + "','" + TARIFF_REFUND_5 + "','" + TARIFF_REFUND + "',"; //관세환급
                            sql += "'" + LEASE_INCOME_1 + "','" + LEASE_INCOME_2 + "' ,'" + LEASE_INCOME_3 + "','" + LEASE_INCOME_4 + "','" + LEASE_INCOME_5 + "','" + LEASE_INCOME + "',"; //임대보증금/임대료 입금
                            sql += "'" + IMPORT_INCOME_1 + "','" + IMPORT_INCOME_2 + "' ,'" + IMPORT_INCOME_3 + "','" + IMPORT_INCOME_4 + "','" + IMPORT_INCOME_5 + "','" + IMPORT_INCOME + "',"; //수입이자입금
                            sql += "'" + ETC_1 + "','" + ETC_2 + "' ,'" + ETC_3 + "','" + ETC_4 + "','" + ETC_5 + "','" + ETC + "',"; //기타
                            sql += "'" + BUSINESS_INCOME + "',"; //[기타] 영업활동상의 자급수입
                            sql += "'" + IV_1 + "','" + IV_2 + "' ,'" + IV_3 + "','" + IV_4 + "','" + IV_5 + "','" + IV + "',"; //원자재 매입대금 지급
                            sql += "'" + M_PAY_1 + "','" + M_PAY_2 + "' ,'" + M_PAY_3 + "','" + M_PAY_4 + "','" + M_PAY_5 + "','" + M_PAY + "',"; //급여와 상여
                            sql += "'" + M_RETIRE_1 + "','" + M_RETIRE_2 + "' ,'" + M_RETIRE_3 + "','" + M_RETIRE_4 + "','" + M_RETIRE_5 + "','" + M_RETIRE + "',"; //퇴직급의 지급
                            sql += "'" + M_FOUNTAIN_1 + "','" + M_FOUNTAIN_2 + "' ,'" + M_FOUNTAIN_3 + "','" + M_FOUNTAIN_4 + "','" + M_FOUNTAIN_5 + "','" + M_FOUNTAIN + "',"; //원천제세 납부
                            sql += "'" + M_WELRARE_1 + "','" + M_WELRARE_2 + "' ,'" + M_WELRARE_3 + "','" + M_WELRARE_4 + "','" + M_WELRARE_5 + "','" + M_WELRARE + "',"; //법정복리비 납부
                            sql += "'" + PAYROLL_COSTS_EXPENSE + "',"; //인건비의 지급
                            sql += "'" + PAYROLL_COSTS_EXPENSE_1 + "','" + PAYROLL_COSTS_EXPENSE_2 + "' ,'" + PAYROLL_COSTS_EXPENSE_3 + "','" + PAYROLL_COSTS_EXPENSE_4 + "','" + PAYROLL_COSTS_EXPENSE_5 + "',"; //인건비의 지급
                            sql += "'" + EXPENSE + "','" + EXPENSE_1 + "' ,'" + EXPENSE_2 + "','" + EXPENSE_3 + "','" + EXPENSE_4 + "','" + EXPENSE_5 + "',"; //사업부경비
                            sql += "'" + LEASE_EXPENSE + "','" + LEASE_EXPENSE_1 + "' ,'" + LEASE_EXPENSE_2 + "','" + LEASE_EXPENSE_3 + "','" + LEASE_EXPENSE_4 + "','" + LEASE_EXPENSE_5 + "',"; //임대보증금/임대료지급
                            sql += "'" + INSEREST + "','" + INSEREST_1 + "' ,'" + INSEREST_2 + "','" + INSEREST_3 + "','" + INSEREST_4 + "','" + INSEREST_5 + "',"; //지급이자
                            sql += "'" + BUSINESS_ETC + "','" + BUSINESS_EXPENSE + "' ,'" + BUSINESS_INFLOW + "',"; //영업활동상의 자금수입] 기타,[영업활동상의 자금지출]
                            sql += "'" + FIXED_ASSET_LAND + "','" + FIXED_ASSET_LAND_1 + "' ,'" + FIXED_ASSET_LAND_2 + "','" + FIXED_ASSET_LAND_3 + "' ,'" + FIXED_ASSET_LAND_4 + "','" + FIXED_ASSET_LAND_5 + "', ";
                            sql += "'" + FIXED_ASSET_MACHINE + "','" + FIXED_ASSET_MACHINE_1 + "' ,'" + FIXED_ASSET_MACHINE_2 + "','" + FIXED_ASSET_MACHINE_3 + "','" + FIXED_ASSET_MACHINE_4 + "','" + FIXED_ASSET_MACHINE_5 + "',"; //[고정자산의 매입] 기계장치 취득
                            sql += "'" + FIXED_ASSET_CAR + "','" + FIXED_ASSET_CAR_1 + "' ,'" + FIXED_ASSET_CAR_2 + "','" + FIXED_ASSET_CAR_3 + "' ,'" + FIXED_ASSET_CAR_4 + "','" + FIXED_ASSET_CAR_5 + "',"; //[고정자산의 매입] 차량 및 공기구 취득
                            sql += "'" + FIXED_ASSET_INCOME + "','" + FIXED_ASSET_INCOME_1 + "' ,'" + FIXED_ASSET_INCOME_2 + "','" + FIXED_ASSET_INCOME_3 + "' ,'" + FIXED_ASSET_INCOME_4 + "','" + FIXED_ASSET_INCOME_5 + "',"; //[고정자산의 매입] 보증금등의 회수
                            sql += "'" + FIXED_ASSET_ETC + "','" + FIXED_ASSET_ETC1 + "' ,'" + FIXED_ASSET_ETC2 + "','" + FIXED_ASSET_ETC3 + "' ,'" + FIXED_ASSET_ETC4 + "','" + FIXED_ASSET_ETC5 + "', "; //[고정자산의 매입] 기타
                            sql += "'" + FIXED_ASSET + "','" + LOAN_EXPENSE + "' ,'" + FINANCIAL_OUT + "','" + INVEST_CASH + "',"; //[고정자산의 매입]대여금,장기금융상품의 매입,투자유가증권 / 출자금 지급
                            sql += "'" + BEWBORROW_0 + "','" + BEWBORROW_1 + "' ,'" + BEWBORROW_2 + "','" + BEWBORROW_3 + "' ,'" + BEWBORROW_4 + "','" + BEWBORROW_5 + "','" + BEWBORROW + "',"; // [투자활동의 자금흐름]신규차입
                            sql += "'" + INVEST_USANCE_1 + "','" + INVEST_USANCE_2 + "' ,'" + INVEST_USANCE_3 + "','" + INVEST_USANCE_4 + "' ,'" + INVEST_USANCE_5 + "','" + INVEST_USANCE + "',"; //[투자활동의 자금흐름]USANCE 차입
                            sql += "'" + INCREASE_CASHIN + "',";
                            sql += "'" + LOAN_CD_0 + "','" + LOAN_CD1 + "' ,'" + LOAN_CD2 + "','" + LOAN_CD3 + "' ,'" + LOAN_CD4 + "','" + LOAN_CD5 + "','" + LOAN_CD + "', "; //차입금의 상환
                            sql += "'" + USANCE_CD1 + "','" + USANCE_CD2 + "' ,'" + USANCE_CD3 + "','" + USANCE_CD4 + "' ,'" + USANCE_CD5 + "','" + USANCE_CD + "',"; //
                            sql += "'" + DIVIDEND_CASHOUT + "',"; //배당금/자기주식 취득 외
                            sql += "'" + CASH_OUT + "','" + CASH_IN + "' ,'" + LAST_AMT + "','" + LAST_LESS_AMT + "' ,'" + AFTER_AMT + "' ,'',"; //재무적지출~비고
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
            //ReportViewer1.Reset();
            grid_regist_excel.DataSource = null;
            grid_regist_excel.DataBind();
        }


        /******************************* 조회 *************************************************************************************/


        protected void btn_select_Click(object sender, EventArgs e) //조회 버튼
        {
            if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
            {
                MessageBox.ShowMessage("'년도'를 입력하세요.", this.Page);

                return;
            }

            if (txt_mm == null || txt_mm.Text.Equals(""))
            {
                MessageBox.ShowMessage("'월'을 입력하세요.", this.Page);

                return;
            }

            if (txt_dd == null || txt_dd.Text.Equals(""))
            {
                MessageBox.ShowMessage("'일'을 입력하세요.", this.Page);

                return;
            }

            ReportViewer1.Reset();

            string sql = "SELECT mm,dd,LENDER_COLLET_1,LENDER_COLLET_2,LENDER_COLLET_3,LENDER_COLLET_4,LENDER_COLLET_5";
            sql += " ,LENDER_COLLET,SURTAX_REFUND_1,SURTAX_REFUND_2,SURTAX_REFUND_3,SURTAX_REFUND_4,SURTAX_REFUND_5,SURTAX_REFUND";
            sql += " ,TARIFF_REFUND_1,TARIFF_REFUND_2,TARIFF_REFUND_3,TARIFF_REFUND_4,TARIFF_REFUND_5,TARIFF_REFUND ";
            sql += " ,LEASE_INCOME_1,LEASE_INCOME_2,LEASE_INCOME_3,LEASE_INCOME_4,LEASE_INCOME_5,LEASE_INCOME";
            sql += " ,IMPORT_INCOME_1,IMPORT_INCOME_2,IMPORT_INCOME_3,IMPORT_INCOME_4,IMPORT_INCOME_5,IMPORT_INCOME";
            sql += " ,ETC_1,ETC_2,ETC_3,ETC_4,ETC_5,ETC";
            sql += " ,BUSINESS_INCOME";
            sql += " ,IV_1,IV_2,IV_3,IV_4,IV_5,IV";
            sql += " ,M_PAY_1,M_PAY_2,M_PAY_3,M_PAY_4,M_PAY_5,M_PAY";
            sql += " ,M_RETIRE_1,M_RETIRE_2,M_RETIRE_3,M_RETIRE_4,M_RETIRE_5,M_RETIRE ";
            sql += " ,M_FOUNTAIN_1,M_FOUNTAIN_2,M_FOUNTAIN_3,M_FOUNTAIN_4,M_FOUNTAIN_5,M_FOUNTAIN ";
            sql += " ,M_WELRARE_1,M_WELRARE_2,M_WELRARE_3,M_WELRARE_4,M_WELRARE_5,M_WELRARE";
            sql += " ,PAYROLL_COSTS_EXPENSE,PAYROLL_COSTS_EXPENSE_1,PAYROLL_COSTS_EXPENSE_2,PAYROLL_COSTS_EXPENSE_3,PAYROLL_COSTS_EXPENSE_4,PAYROLL_COSTS_EXPENSE_5,EXPENSE";
            sql += " ,EXPENSE_1,EXPENSE_2,EXPENSE_3,EXPENSE_4,EXPENSE_5,LEASE_EXPENSE";
            sql += " ,LEASE_EXPENSE_1,LEASE_EXPENSE_2,LEASE_EXPENSE_3,LEASE_EXPENSE_4,LEASE_EXPENSE_5";
            sql += " ,INSEREST,INSEREST_1,INSEREST_2,INSEREST_3,INSEREST_4,INSEREST_5";
            sql += " ,BUSINESS_ETC,BUSINESS_EXPENSE,BUSINESS_INFLOW ";
            sql += " ,FIXED_ASSET_LAND,FIXED_ASSET_LAND_1,FIXED_ASSET_LAND_2,FIXED_ASSET_LAND_3,FIXED_ASSET_LAND_4,FIXED_ASSET_LAND_5 ";
            sql += " ,FIXED_ASSET_MACHINE,FIXED_ASSET_MACHINE_1,FIXED_ASSET_MACHINE_2,FIXED_ASSET_MACHINE_3,FIXED_ASSET_MACHINE_4,FIXED_ASSET_MACHINE_5 ";
            sql += " ,FIXED_ASSET_CAR,FIXED_ASSET_CAR_1,FIXED_ASSET_CAR_2,FIXED_ASSET_CAR_3,FIXED_ASSET_CAR_4,FIXED_ASSET_CAR_5";
            sql += " ,FIXED_ASSET_INCOME,FIXED_ASSET_INCOME_1,FIXED_ASSET_INCOME_2,FIXED_ASSET_INCOME_3,FIXED_ASSET_INCOME_4,FIXED_ASSET_INCOME_5";
            sql += " ,FIXED_ASSET_ETC,FIXED_ASSET_ETC1,FIXED_ASSET_ETC2,FIXED_ASSET_ETC3,FIXED_ASSET_ETC4,FIXED_ASSET_ETC5";
            sql += " ,FIXED_ASSET,LOAN_EXPENSE,FINANCIAL_OUT,INVEST_CASH";
            sql += " ,BEWBORROW_0,BEWBORROW_1,BEWBORROW_2,BEWBORROW_3,BEWBORROW_4,BEWBORROW_5,BEWBORROW";
            sql += " ,INVEST_USANCE_1,INVEST_USANCE_2,INVEST_USANCE_3,INVEST_USANCE_4,INVEST_USANCE_5,INVEST_USANCE";
            sql += " ,INCREASE_CASHIN";
            sql += " ,LOAN_CD_0,LOAN_CD1,LOAN_CD2,LOAN_CD3,LOAN_CD4,LOAN_CD5,LOAN_CD ";
            sql += " ,USANCE_CD1,USANCE_CD2,USANCE_CD3,USANCE_CD4,USANCE_CD5,USANCE_CD  ";
            sql += " ,DIVIDEND_CASHOUT,CASH_OUT,CASH_IN,LAST_AMT,LAST_LESS_AMT,AFTER_AMT";
            sql += " FROM a_daily_amt ";
            sql += "WHERE YYYY = '" + txt_yyyy.Text.Trim() + "' AND MM = '" + txt_mm.Text.Trim() + "' ";
            sql += "and DD = '" + txt_dd.Text.Trim() + "' ";

            ReportViewer1.Reset();
            ds_am_aa1001 dt1 = new ds_am_aa1001();
            ReportCreator(dt1, sql, ReportViewer1, "rp_am_aa1001.rdlc", "DataSet1");
        }

        /******************************* 조회 *************************************************************************************/


        protected void btn_save2_Click(object sender, EventArgs e) //저장2
        {

            if
               (txt_yyyy1.Text == "" || txt_mm1.Text == "" || tb_amt1.Text == "" ||
                tb_amt2.Text == "" || tb_amt3.Text == "" || tb_amt4.Text == "" ||
                tb_amt5.Text == "" || tb_amt6.Text == "" || tb_amt7.Text == "" ||
                tb_amt8.Text == "" || tb_amt9.Text == "" || tb_amt10.Text == "" ||
                tb_amt11.Text == "" || tb_amt12.Text == "" || tb_amt13.Text == "" ||
                tb_amt14.Text == "" || tb_amt16.Text == "" || tb_amt17.Text == "" ||
                tb_amt18.Text == "" || tb_amt19.Text == "" || tb_amt21.Text == "" ||
                tb_amt22.Text == "" || tb_amt23.Text == "" || tb_amt24.Text == "" ||
                tb_amt26.Text == "" || tb_amt27.Text == "" || tb_amt28.Text == "" ||
                tb_amt29.Text == "")
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력되지 않은 항목이 있습니다.')", true);
            }

            else
            {
                conn_erp.Open();
                cmd_erp.Connection = conn_erp;
                SqlTransaction tran_insert = conn_erp.BeginTransaction();
                cmd_erp.Transaction = tran_insert;

                try
                {

                    string insert_queryStr;
                    insert_queryStr = "INSERT INTO A_DAILY_AMT_ETC VALUES('" + txt_yyyy1.Text + "',";
                    insert_queryStr += "'" + txt_mm1.Text + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt1.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt2.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt3.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt4.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt5.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt6.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt7.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt8.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt9.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt10.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt11.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt12.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt13.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt14.Text) + "',";
                    insert_queryStr += "'" + (Convert.ToDouble(tb_amt11.Text) + Convert.ToDouble(tb_amt12.Text) + Convert.ToDouble(tb_amt13.Text) + Convert.ToDouble(tb_amt14.Text)).ToString() + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt16.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt17.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt18.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt19.Text) + "',";
                    insert_queryStr += "'" + (Convert.ToDouble(tb_amt16.Text) + Convert.ToDouble(tb_amt17.Text) + Convert.ToDouble(tb_amt18.Text) + Convert.ToDouble(tb_amt19.Text)).ToString() + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt21.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt22.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt23.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt24.Text) + "',";
                    insert_queryStr += "'" + (Convert.ToDouble(tb_amt21.Text) + Convert.ToDouble(tb_amt22.Text) + Convert.ToDouble(tb_amt23.Text) + Convert.ToDouble(tb_amt24.Text)).ToString() + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt26.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt27.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt28.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt29.Text) + "',";
                    insert_queryStr += "'" + (Convert.ToDouble(tb_amt26.Text) + Convert.ToDouble(tb_amt27.Text) + Convert.ToDouble(tb_amt28.Text) + Convert.ToDouble(tb_amt29.Text)).ToString() + "',";
                    insert_queryStr += "'0',";
                    insert_queryStr += "'0',";
                    insert_queryStr += "'" + Session["User"].ToString() + "',";
                    insert_queryStr += "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                    insert_queryStr += "'" + Session["User"].ToString() + "',";
                    insert_queryStr += "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                    cmd_erp.CommandText = insert_queryStr;
                    cmd_erp.ExecuteNonQuery();
                    tran_insert.Commit();

                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('저장되었습니다.')", true);

                }

                catch
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜에 저장된 데이터가 존재합니다.')", true);
                    tran_insert.Rollback();
                }
            }
            conn_erp.Close();

        }

        /******************************* 조회 *************************************************************************************/

        protected void btn_select2_Click(object sender, EventArgs e) //조회2
        {
            string[] temp;

            if (txt_yyyy1 == null || txt_yyyy1.Text.Equals(""))
            {
                MessageBox.ShowMessage("'년도'를 입력하세요.", this.Page);

                return;
            }

            if (txt_mm1 == null || txt_mm1.Text.Equals(""))
            {
                MessageBox.ShowMessage("'월'을 입력하세요.", this.Page);

                return;
            }
            string queryStr = "SELECT * FROM A_DAILY_AMT_ETC WHERE YYYY = '" + txt_yyyy1.Text + "' and  MM = '" + txt_mm1.Text + "'";

          
            cmd_erp.Connection = conn_erp;
            conn_erp.Open();
            cmd_erp.CommandText = queryStr;
            SqlDataReader dReader_select = cmd_erp.ExecuteReader();

            temp = new string[dReader_select.FieldCount];

            if (dReader_select.Read())
            {
                for (int a = 0; a < dReader_select.FieldCount; a++)
                {
                    temp[a] = dReader_select[a].ToString();
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜의 데이터가 없습니다.')", true);
            }

            txt_yyyy1.Text = temp[0];
            txt_mm1.Text = temp[1];
            tb_amt1.Text = temp[2];
            tb_amt2.Text = temp[3];
            tb_amt3.Text = temp[4];
            tb_amt4.Text = temp[5];
            tb_amt5.Text = temp[6];
            tb_amt6.Text = temp[7];
            tb_amt7.Text = temp[8];
            tb_amt8.Text = temp[9];
            tb_amt9.Text = temp[10];
            tb_amt10.Text = temp[11];
            tb_amt11.Text = temp[12];
            tb_amt12.Text = temp[13];
            tb_amt13.Text = temp[14];
            tb_amt14.Text = temp[15];
            tb_amt16.Text = temp[17];
            tb_amt17.Text = temp[18];
            tb_amt18.Text = temp[19];
            tb_amt19.Text = temp[20];
            tb_amt21.Text = temp[22];
            tb_amt22.Text = temp[23];
            tb_amt23.Text = temp[24];
            tb_amt24.Text = temp[25];
            tb_amt26.Text = temp[27];
            tb_amt27.Text = temp[28];
            tb_amt28.Text = temp[29];
            tb_amt29.Text = temp[30];


            btn_save2.Enabled = false;

            conn_erp.Close();
            dReader_select.Close();
            conn_erp.Dispose();
        }
        /******************************* 삭제 *************************************************************************************/
        protected void btn_delete2_Click(object sender, EventArgs e) //삭제 버튼
        {
           conn_erp.Open();

            string delete_queryStr;
            SqlTransaction tran_delete = conn_erp.BeginTransaction();
            cmd_erp.Connection = conn_erp;
            cmd_erp.Transaction = tran_delete;

            int intReturnRow_DELETE;

            try
            {
                delete_queryStr = "DELETE A_DAILY_AMT_ETC WHERE YYYY = '" + txt_yyyy1.Text + "' AND MM = '" + txt_mm1.Text + "'";
                cmd_erp.CommandText = delete_queryStr;
                intReturnRow_DELETE = cmd_erp.ExecuteNonQuery();


                if (intReturnRow_DELETE == 0)
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('삭제할 데이터가 없습니다.')", true);
                }

                else
                {
                    tran_delete.Commit();
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('삭제되었습니다.')", true);
                }
            }
            catch
            {
                tran_delete.Rollback();
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('에러가 발생되었습니다. 다시 시도하세요.')", true);
            }

            conn_erp.Close();
        }
        }


    }
