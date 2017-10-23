using System;
using System.Data;
using System.Text;
//using System.Data.SqlClient;
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
using ERPAppAddition.QueryExe;
using Microsoft.Reporting.WebForms;
using FarPoint.Web.Spread.Data;

/*
 
 */
namespace ERPAppAddition.ERPAddition.SM.sm_s3001
{
    public partial class sm_s3001 : System.Web.UI.Page
    {
        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;
        OracleDataAdapter sqlAdapter1;

        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        SqlDataAdapter erp_sqlAdapter;
        DataSet ds = new DataSet();
        cls_dbexe_erp dbexe = new cls_dbexe_erp();

        string[,] array = new string[30, 30]; //new invoice를 저장할 배열
        string[] cbstr;
        string userid, db_name;
        FarPoint.Web.Spread.ComboBoxCellType combo_cell;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                //초기화면설정
                rbtnl_chk_process_SelectedIndexChanged(null, null);
                //거래처 설정
                FillDropDownList();

                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;
                //FarPoint.Web.Spread.CheckBoxCellType check_cell = new FarPoint.Web.Spread.CheckBoxCellType();
                //FpSpread_new_data.ActiveSheetView.Columns[0].CellType = check_cell;
                //체크박스 전체선택용 함수 호출
                myCheck c1 = new myCheck();
                FpSpread_new_data.ActiveSheetView.ColumnHeader.Cells[0, 0].CellType = c1;
                for (int i = 0; i < FpSpread_new_data.ActiveSheetView.RowCount; i++)
                    FpSpread_new_data.ActiveSheetView.Cells[i, 0].CellType = c1;

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

        private void FillDropDownList()
        {
            string sql = "SELECT SYSCODE_NAME cust_cd FROM SYSCODEDATA A WHERE  A.PLANT = 'CCUBEDIGITAL' AND A.SYSTABLE_NAME IN ( 'CUSTOMER') UNION ALL SELECT '%' FROM DUAL  ORDER BY 1 ";
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            try
            {
                // 품목 드랍다운리스트 내용을 보여준다.
                OracleCommand cmd2 = new OracleCommand(sql, conn);

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

        protected void rbtnl_chk_process_SelectedIndexChanged(object sender, EventArgs e)
        {
            //각 컨트롤값 초기화하기
            tb_new_invoice_no.Text = "";
            tb_invoice_no.Text = "";
            tb_ship_fr_cust_nm.Text = "";
            hf_tb_ship_fr_cust_cd.Value = null;  //발신인 거래처코드
            tb_ship_fr_cust_nm.Text = "";
            tb_ship_to_add.Text = "";                         //발신인주소
            tb_ship_to_tel.Text = "";                         //발신인전화번호
            tb_ship_fr_fax.Text = "";                        //발신인 fax
            hf_tb_bill_to_cust_cd.Value = null;  //수취인 거래처코드
            tb_bill_to_cust_nm.Text = "";
            tb_bill_to_add.Text = "";                          //수취인 주소
            tb_bill_to_tel.Text = "";                         //수취인 전화번호
            tb_bill_to_fax.Text = "";                         //수취인 fax
            tb_bill_to_name.Text = "";                       //수취인
            hf_tb_ship_to_cust_cd.Value = null;  //실물수령인 거래처코드
            tb_ship_to_cust_nm.Text = "";
            tb_ship_to_add.Text = "";                       //실물수령인 주소
            tb_ship_to_tel.Text = "";                         //실물수령인 전화
            tb_ship_to_fax.Text = "";                        //실물수령인 fax
            tb_ship_to_name.Text = "";                        //실물수령인 
            tb_port_of_loading.Text = "";                 //출발지
            tb_final_destination.Text = "";             //도착지
            tb_carrier.Text = "";                             //운송업체
            tb_board_on_about.Text = "";                   //발송일
            tb_invoice_dt.Text = "";                          //인보이스 발행일
            tb_remark.Text = "";                                   //비고
            tb_remark_incoterms.Text = "";               //운임조건
            tb_remark_pay_type.Text = "";                 //유무상구분
            tb_total_box_cnt.Text = "";                      //전체박스 수량
            tb_hts_code.Text = "";                               //hts code
            tb_country_of_org.Text = "";                    //원산지
            tb_bank_info.Text = "";                               //회사 은행 정보
            tb_net_weight.Text = "";                           //net weight
            tb_gross_weight.Text = "";                       //gross weight
            tb_bank_name.Text = "";
            tb_bill_to_add.Text = "";
            tb_bank_addr.Text = "";
            tb_bank_branch.Text = "";
            tb_bank_swiftcode.Text = "";
            tb_bank_acct_no.Text = "";
            tb_bank_accountee.Text = "";
            //하단 mes 데이타선택 조회부분 초기화
            tb_fr_yyyymmdd.Text = "";
            tb_to_yyyymmdd.Text = "";
            ddl_cust_cd.SelectedValue = "%";

            //spread 초기화
            FpSpread_view_data.DataSource = null;
            FpSpread_view_data.DataBind();
            FpSpread_new_data.DataSource = null;
            FpSpread_new_data.DataBind();

            if (rbtnl_chk_process.SelectedValue == "view")
            {
                Panel_body1.Visible = true;
                Panel_view_data.Visible = true; //조회및 수정용 데이타
                Panel_body2.Visible = false; //신규등록용데이타
                btn_copy.Visible = false;
                Panel_new_invoice_no.Visible = false; //신규인보이스번호등록용

            }
            else
            {
                Panel_body1.Visible = true;
                Panel_view_data.Visible = false;
                Panel_body2.Visible = true;
                btn_copy.Visible = true;
                Panel_new_invoice_no.Visible = true;
            }
        }

        protected void rbtn_menu_new_sub1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void SetSpreadCheckBox()
        {
            for (int i = 0; i <= FpSpread_new_data.Rows.Count - 1; i++)
            {


                FarPoint.Web.Spread.CheckBoxCellType check_cell = new FarPoint.Web.Spread.CheckBoxCellType();
                //check_cell.AutoPostBack = true;                

                FarPoint.Web.Spread.Cell cellobj;
                cellobj = FpSpread_new_data.ActiveSheetView.Cells[i, 0, i, 0];
                cellobj.CellType = check_cell;
            }
        }

        private void SetSpreadDropDown(int column)
        {
            for (int i = 0; i <= FpSpread_new_data.Rows.Count - 1; i++)
            {
                cbstr = new String[] { "", "USD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "RUB", "SGD" };

                combo_cell = new FarPoint.Web.Spread.ComboBoxCellType(cbstr);
                combo_cell.ShowButton = true; //콤보박스버튼보여주기
                combo_cell.UseValue = true;   //
                combo_cell.CssClass = "style2";
                //combo_cell.
                //combo_cell.OnClientChanged = "alert(\'You selected the item\');";
                FarPoint.Web.Spread.Cell cellobj;
                cellobj = FpSpread_new_data.ActiveSheetView.Cells[i, column, i, column];
                cellobj.CellType = combo_cell;
            }
        }

        private void SetSpreadDropDown_view(int column)
        {
            for (int i = 0; i <= FpSpread_view_data.Rows.Count - 1; i++)
            {
                cbstr = new String[] { "", "USD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "RUB", "SGD" };

                combo_cell = new FarPoint.Web.Spread.ComboBoxCellType(cbstr);
                combo_cell.ShowButton = true; //콤보박스버튼보여주기
                combo_cell.UseValue = true;   //
                combo_cell.CssClass = "style2";
                //combo_cell.OnClientChanged = "alert(\'You selected the item\');";
                FarPoint.Web.Spread.Cell cellobj;
                cellobj = FpSpread_view_data.ActiveSheetView.Cells[i, column, i, column];
                cellobj.CellType = combo_cell;
            }
        }
        //생성시 보여지는 spread lock 용
        private void SetSpreadColumnLock(int column)
        {
            FpSpread_new_data.ActiveSheetView.Protect = true;
            FpSpread_new_data.ActiveSheetView.LockBackColor = Color.LightCyan;
            FpSpread_new_data.ActiveSheetView.LockForeColor = Color.Green;

            FarPoint.Web.Spread.Column columnobj;

            int columncnt = FpSpread_new_data.Columns.Count;

            columnobj = FpSpread_new_data.ActiveSheetView.Columns[1, column]; // 입력된칼럼 lock
            columnobj.Locked = true;

        }
        // 수정용시 보여지는 spread lock용
        private void SetSpreadColumnLock_ViewData(int column)
        {
            FpSpread_view_data.ActiveSheetView.Protect = true;
            FpSpread_view_data.ActiveSheetView.LockBackColor = Color.LightCyan;
            FpSpread_view_data.ActiveSheetView.LockForeColor = Color.Green;

            FarPoint.Web.Spread.Column columnobj;

            int columncnt = FpSpread_view_data.Columns.Count;

            columnobj = FpSpread_view_data.ActiveSheetView.Columns[0, column]; // 입력된칼럼 lock
            columnobj.Locked = true;

        }

        protected void btn_mighty_retrieve_Click(object sender, EventArgs e)
        {
            string fr_dt, to_dt, cust_nm;
            fr_dt = tb_fr_yyyymmdd.Text;
            to_dt = tb_to_yyyymmdd.Text;
            if (fr_dt == to_dt)
                to_dt = DateTime.ParseExact(to_dt, "yyyyMMdd", null).AddMonths(0).AddDays(1).ToString("yyyyMMdd");

            SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
            SqlCommand cmd_erp = new SqlCommand();
            SqlDataReader dr_erp;
            SqlDataAdapter erp_sqlAdapter;


            cust_nm = ddl_cust_cd.SelectedValue;

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();

            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.USP_SM_S3001_MES_INVOICE_VIEW";
            cmd_erp.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@cust_nm", SqlDbType.VarChar, 30);
            SqlParameter param2 = new SqlParameter("@fr_dt", SqlDbType.VarChar, 20);
            SqlParameter param3 = new SqlParameter("@to_dt", SqlDbType.VarChar, 20);

            param1.Value = cust_nm;
            param2.Value = fr_dt;
            param3.Value = to_dt;

            cmd_erp.Parameters.Add(param1);
            cmd_erp.Parameters.Add(param2);
            cmd_erp.Parameters.Add(param3);

            try
            {
                
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                DataTable dt = new DataTable();
                                
                da.Fill(dt);
                FpSpread_new_data.DataSource = dt;
                //첫칼럼은 체크박스로 사용하기 위해 AutoGenerateColumns = false 로 지정후 일일이 칼럼을 지정해줌
                FpSpread_new_data.Columns[1].DataField = "인보이스번호";
                FpSpread_new_data.Columns[2].DataField = "거래처명";
                FpSpread_new_data.Columns[3].DataField = "패킹리스트날짜";
                FpSpread_new_data.Columns[4].DataField = "패킹리스트번호";
                FpSpread_new_data.Columns[5].DataField = "패킹타입";
                FpSpread_new_data.Columns[6].DataField = "디바이스명";
                FpSpread_new_data.Columns[7].DataField = "고객LOT";
                FpSpread_new_data.Columns[8].DataField = "LOT";
                FpSpread_new_data.Columns[9].DataField = "수량";
                FpSpread_new_data.Columns[10].DataField = "LOT단위";
                FpSpread_new_data.Columns[11].DataField = "PO번호";
                FpSpread_new_data.Columns[12].DataField = "FABLOTNO";
                FpSpread_new_data.Columns[13].DataField = "COM_NONC";
                FpSpread_new_data.Columns[14].DataField = "단가화폐";
                FpSpread_new_data.Columns[15].DataField = "단가";
                FpSpread_new_data.Columns[16].DataField = "COM_NONC";
                FpSpread_new_data.Columns[17].DataField = "원자재화폐";
                FpSpread_new_data.Columns[18].DataField = "원자재단가";
                FpSpread_new_data.Columns[19].DataField = "COM_NONC";
                FpSpread_new_data.Columns[20].DataField = "제3출하처화폐";
                FpSpread_new_data.Columns[21].DataField = "제3출하처단가";
                FpSpread_new_data.DataBind();

                //SetSpreadCheckBox();
                //SetSpreadDropDown(13); //단가화폐
                //SetSpreadDropDown(15); //원자재화폐
                //SetSpreadDropDown(17); //제3출하처화폐
                //SetSpreadColumnLock(12);
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }

        }

        protected void btn_pop_ship_fr_cust_cd_Click(object sender, EventArgs e)
        {
            Response.Write("<script>window.open('../../POPUP/pop_cust_cd.aspx?pgid=sm_s3001&popupid=1','','top=100,left=100,width=800,height=600')</script>");

        }

        protected void btn_pop_bill_cust_cd_Click(object sender, EventArgs e)
        {
            Response.Write("<script>window.open('../../POPUP/pop_cust_cd.aspx?pgid=sm_s3001&popupid=2','','top=100,left=100,width=800,height=600')</script>");
        }

        protected void btn_pop_ship_to_cust_cd_Click(object sender, EventArgs e)
        {
            Response.Write("<script>window.open('../../POPUP/pop_cust_cd.aspx?pgid=sm_s3001&popupid=3','','top=100,left=100,width=800,height=600')</script>");
        }

        protected void btn_mighty_save_Click(object sender, EventArgs e)
        {
            FpSpread_new_data.SaveChanges();
            //FpSpread_new_data_UpdateCommand(null, null);
            btn_mighty_retrieve_Click(null, null);
        }
        protected void btn_pop_invoice_Click(object sender, EventArgs e)
        {
            Response.Write("<script>window.open('pop_sm_s3001_invoice.aspx?pgid=sm_s3001&popupid=1','','top=100,left=100,width=800,height=600')</script>");

        }
        protected void FpSpread_new_data_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e) 
        {
            int colcnt, chk_exist_data = 0;
            int i, s_column = 0, j = 0;
            int r = (int)e.CommandArgument;

            string a_price_unit = "", b_price_unit = "", c_price_unit = "";
            string a_price = "", b_price = "", c_price = "";
            string a_type ="", b_type="", c_type="";  //단가, 원자재, 제3출하처 타입 추가
            string invoice_no, customer, issue_time, packinglist_no, packing_type, part;
            string cust_lot, lot_number, lot_qty, lot_unit, po_no, fablotno;
            //colcnt = e.EditValues.Count - 1;
            //colcnt = FpSpread_new_data.Rows.Count - 1;

            string b = e.EditValues[0].ToString() ;
            //체크된 로직인지 확인하기
            if (e.EditValues[0].ToString() == "True") //체크된 row이면
            //for (int r = 0; r <= FpSpread_new_data.Rows.Count - 1; r++) //row 수만큼 루프
            {
                //string chk = FpSpread_new_data.ActiveSheetView.Cells[r, 0].Text;
                //if (chk == "True")
                {
                    //조회된 값을 가져온다.
                    invoice_no = tb_invoice_no.Text; //FpSpread_new_data.Sheets[0].Cells[r, 1].Value.ToString();
                    if (FpSpread_new_data.Sheets[0].Cells[r, 2].Value == null)
                        customer = "";
                    else
                        customer = FpSpread_new_data.Sheets[0].Cells[r, 2].Value.ToString();
                    if (FpSpread_new_data.Sheets[0].Cells[r, 3].Value == null)
                        issue_time = "";
                    else
                        issue_time = FpSpread_new_data.Sheets[0].Cells[r, 3].Value.ToString();
                    if (FpSpread_new_data.Sheets[0].Cells[r, 4].Value == null)
                        packinglist_no = "";
                    else
                        packinglist_no = FpSpread_new_data.Sheets[0].Cells[r, 4].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 5].Value == null)
                        packing_type = "";
                    else
                        packing_type = FpSpread_new_data.Sheets[0].Cells[r, 5].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 6].Value == null)
                        part = "";
                    else
                        part = FpSpread_new_data.Sheets[0].Cells[r, 6].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 7].Value == null)
                        cust_lot = "";
                    else
                        cust_lot = FpSpread_new_data.Sheets[0].Cells[r, 7].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 8].Value == null)
                        lot_number = "";
                    else
                        lot_number = FpSpread_new_data.Sheets[0].Cells[r, 8].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 9].Value == null)
                        lot_qty = "";
                    else
                        lot_qty = FpSpread_new_data.Sheets[0].Cells[r, 9].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 10].Value == null)    //lot 단위
                        lot_unit = "";
                    else
                        lot_unit = FpSpread_new_data.Sheets[0].Cells[r, 10].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 11].Value == null)  //po 번호
                        po_no = "";
                    else
                        po_no = FpSpread_new_data.Sheets[0].Cells[r, 11].Value.ToString();

                    if (FpSpread_new_data.Sheets[0].Cells[r, 12].Value == null)  //fablotno
                        fablotno = "";
                    else
                        fablotno = FpSpread_new_data.Sheets[0].Cells[r, 12].Value.ToString();

                    if (invoice_no == "" || invoice_no == null)
                    {
                        MessageBox.ShowMessage("상단 인보이스번호를 입력해주세요. ", this.Page);
                        tb_invoice_no.Focus();
                    }
                    else
                    {
                        if (e != null) //클릭해서 수정한 건이 있으면
                        {
                            a_type = e.EditValues[13].ToString(); // 단가 구분 추가
                            a_price_unit = e.EditValues[14].ToString(); //13->14 수정
                            a_price = e.EditValues[15].ToString();//14->15 수정

                            b_type = e.EditValues[16].ToString(); // 원자재단가 구분 추가
                            b_price_unit = e.EditValues[17].ToString();//15->17
                            b_price = e.EditValues[18].ToString();//16->18

                            c_type = e.EditValues[19].ToString(); //제3출하처 구분 추가
                            c_price_unit = e.EditValues[20].ToString(); //17->20
                            c_price = e.EditValues[21].ToString(); //18->21
                        }
                        if (a_type == null || a_type == "" || a_type == "System.Object") //COST_TYPE
                        {
                            a_type = "";
                        }
                        if (a_price_unit == null || a_price_unit == "" || a_price_unit == "System.Object") //COST_CUR_CD
                        {
                            a_price_unit = "";
                        }
                        if (a_price == null || a_price == "" || a_price == "System.Object") //COST_PRICE
                        {
                            a_price = "0";
                        }

                        if (b_type == null || b_type == "" || b_type == "System.Object") //MAT_TYPE
                        {
                            b_type = "";
                        }
                        if (b_price_unit == null || b_price_unit == "" || b_price_unit == "System.Object") //MAT_CUR_CD
                        {
                            b_price_unit = "";
                        }
                        if (b_price == null || b_price == "" || b_price == "System.Object") //MAT_PRICE
                        {
                            b_price = "0";
                        }
                        if (c_type == null || c_type == "" || c_type == "System.Object") //ATTACH_TYPE
                        {
                            c_type = "";
                        }
                        if (c_price_unit == null || c_price_unit == "" || c_price_unit == "System.Object") //ATTACH_CUR_CD
                        {
                            c_price_unit = "";
                        }
                        if (c_price == null || c_price == "" || c_price == "System.Object") //ATTACH_PRICE
                        {
                            c_price = "0";
                        }

                        // 저장된 이력이 있는지 확인한다. 저장된 값이 없다면 
                        string sql;
                        sql = "select count(*) from SM_INVOICE_HDR_NEPES where invoice_no = '" + tb_invoice_no.Text + "' ";

                        if (dbexe.QueryExecute(sql, "check") > 0) //기존 header에 저장된 값이 있다면...
                        {
                            sql = "select count(*) from SM_INVOICE_DTL_NEPES where invoice_no = '" + tb_invoice_no.Text + "' and ITEM_NM = '" + part + "' and LOT_NO = '" + lot_number + "' and PL_NO = '" + packinglist_no + "' ";
                            if (dbexe.QueryExecute(sql, "check") > 0) //기존 DTL에 저장된 값이 있다면 업데이트, 
                            {
                                sql = " update SM_INVOICE_DTL_NEPES  " +
                                      "    set COST_TYPE = '" + a_type + "' " +  // 단가 구분 추가
                                      "        COST_CUR_CD = '" + a_price_unit + "' " +
                                      "        COST_PRICE = '" + a_price + "' " +
                                      "        MAT_TYPE = '" + b_type + "' " + // 원자재 구분 추가
                                      "        MAT_CUR_CD = '" + b_price_unit + "' " +
                                      "        MAT_PRICE = '" + b_price + "' " +
                                      "        ATTACH_TYPE = '" + c_type + "' " + //제3출하처 구분 추가
                                      "        ATTACH_CUR_CD = '" + c_price_unit + "' " +
                                      "        ATTACH_PRICE  = '" + c_price + "' " +
                                      "        UPDT_USER_ID  = '" + Session["User"].ToString() + "' " +
                                      "        UPDT_DT  = getdate() " +
                                      "  where INVOICE_NO = '" + tb_invoice_no.Text + "'";
                                if (dbexe.QueryExecute(sql, "") > 0) //기존 header에 저장된 값이 있다면...
                                {
                                    MessageBox.ShowMessage("저장되었습니다. ", this.Page);
                                }
                            }
                            else //없다면 insert
                            {

                                sql = " insert into SM_INVOICE_DTL_NEPES values ( " +
                                      "   '" + tb_invoice_no.Text + "' " +
                                      " , '" + customer + "' " +                        // CUST_NM
                                      " , '" + issue_time + "' " +                        // PL_DT
                                      " , '" + packinglist_no + "' " +                        // PL_NO
                                      " , '" + packing_type + "' " +                        // PL_TYPE
                                      " , '" + part + "' " +                        // ITEM_NM
                                      " , '" + cust_lot + "' " +                        // CUST_LOT_NO
                                      " , '" + lot_number + "' " +                        // LOT_NO
                                      " , '" + Convert.ToDecimal(lot_qty) + "' " +       // LOT_QTY
                                      " , '" + lot_unit + "' " +                       // LOT_UNIT
                                      " , '" + po_no + "' " +                       // PO_NO
                                      " , '" + fablotno + "' " +                       // FAB_LOT_NO
                                      " , '" + a_type + "' " +                       // COST_TYPE
                                      " , '" + a_price_unit + "' " +                       // COST_CUR_CD
                                      " , '" + Convert.ToDecimal(a_price) + "' " +      // COST_PRICE
                                      " , '" + b_type + "' " +                       // MAT_TYPE
                                      " , '" + b_price_unit + "' " +                       // MAT_CUR_CD
                                      " , '" + Convert.ToDecimal(b_price) + "' " +      // MAT_PRICE
                                      " , '" + c_type + "' " +                       // ATTACH_TYPE
                                      " , '" + c_price_unit + "' " +                       // ATTACH_CUR_CD
                                      " , '" + Convert.ToDecimal(c_price) + "' " +      // ATTACH_PRICE
                                      " , '" + Session["User"].ToString() + "' " +
                                      " , getdate() " +
                                      " , '" + Session["User"].ToString() + "' " +
                                      " , getdate() " +
                                      " ) ";
                                if (dbexe.QueryExecute(sql, "") > 0) //기존 header에 저장된 값이 있다면...
                                {
                                    MessageBox.ShowMessage("저장되었습니다. ", this.Page);
                                }
                                else
                                {
                                    MessageBox.ShowMessage("저장에 실패했습니다. 관리자에게 문의하세요. ", this.Page);
                                }
                            }
                        }

                    }

                }
            }


        }
        private void fn_save(int r)
        {

        }



        protected void btn_update_Click(object sender, EventArgs e)
        {
            string invoice_no = "", invoice_type1, invoice_type2;
            string ship_fr_cust_cd, ship_fr_cust_nm, ship_fr_add, ship_fr_tel, ship_fr_fax;
            string bill_fr_cust_cd, bill_fr_cust_nm, bill_to_add, bill_to_tel, bill_to_fax, bill_to_name;
            string ship_to_cust_cd, ship_to_cust_nm, ship_to_add, ship_to_tel, ship_to_fax, ship_to_name;
            string port_of_loading, final_destination, carrier, board_on_about;
            string invoice_dt, lc_no_dt, lc_issue_bank, remark, remark_incoterms;
            string remark_pay_type, total_box_cnt, hts_code, country_of_org, bank_cd, net_weight, gross_weight;
            string bank_name, bank_addr, bank_branch, bank_swiftcode, bank_acct_no, bank_accountee;
            string sql = "";
            if (rbtnl_chk_process.SelectedValue == "new") //신규생성이면
                invoice_no = tb_new_invoice_no.Text;                              //인보이스번호
            if (rbtnl_chk_process.SelectedValue == "view") //수정성이면
                invoice_no = tb_invoice_no.Text;                              //인보이스번호
            invoice_type1 = rbtnl_pay_type.SelectedValue;              //인보이스타입1(유상,무상,pra)
            invoice_type2 = rbtnl_target_type.SelectedValue;           //인보이스타입2(고객,면허)
            ship_fr_cust_cd = hf_tb_ship_fr_cust_cd.Value.ToString();  //발신인 거래처코드
            ship_fr_cust_nm = tb_ship_fr_cust_nm.Text;
            ship_fr_add = tb_ship_fr_add.Text;                         //발신인주소
            ship_fr_tel = tb_ship_fr_tel.Text;                         //발신인전화번호
            ship_fr_fax = tb_ship_fr_fax.Text;                         //발신인 fax
            bill_fr_cust_cd = hf_tb_bill_to_cust_cd.Value.ToString();  //수취인 거래처코드
            bill_fr_cust_nm = tb_bill_to_cust_nm.Text;
            bill_to_add = tb_bill_to_add.Text;                         //수취인 주소
            bill_to_tel = tb_bill_to_tel.Text;                         //수취인 전화번호
            bill_to_fax = tb_bill_to_fax.Text;                         //수취인 fax
            bill_to_name = tb_bill_to_name.Text;                         //수취인
            ship_to_cust_cd = hf_tb_ship_to_cust_cd.Value.ToString();  //실물수령인 거래처코드
            ship_to_cust_nm = tb_ship_to_cust_nm.Text;
            ship_to_add = tb_ship_to_add.Text;                         //실물수령인 주소
            ship_to_tel = tb_ship_to_tel.Text;                         //실물수령인 전화
            ship_to_fax = tb_ship_to_fax.Text;                         //실물수령인 fax
            ship_to_name = tb_ship_to_name.Text;                         //실물수령인 
            port_of_loading = tb_port_of_loading.Text;                 //출발지
            final_destination = tb_final_destination.Text;             //도착지
            carrier = tb_carrier.Text;                                 //운송업체
            board_on_about = tb_board_on_about.Text;                   //발송일
            invoice_dt = tb_invoice_dt.Text;                           //인보이스 발행일
            lc_no_dt = null;                                           //신용장번호및발행일
            lc_issue_bank = null;                                      //시용장개설은행
            remark = tb_remark.Text;                                   //비고
            remark_incoterms = tb_remark_incoterms.Text;               //운임조건
            remark_pay_type = tb_remark_pay_type.Text;                 //유무상구분
            total_box_cnt = tb_total_box_cnt.Text;                     //전체박스 수량
            hts_code = tb_hts_code.Text;                               //hts code
            country_of_org = tb_country_of_org.Text;                   //원산지
            bank_cd = hf_tb_bank_cd.Value.ToString();                              //회사 은행 정보
            net_weight = tb_net_weight.Text;                           //net weight
            gross_weight = tb_gross_weight.Text;                       //gross weight
            bank_name = tb_bank_name.Text;
            bank_addr = tb_bank_addr.Text;
            bank_branch = tb_bank_branch.Text;
            bank_swiftcode = tb_bank_swiftcode.Text;
            bank_acct_no = tb_bank_acct_no.Text;
            bank_accountee = tb_bank_accountee.Text;


            // 인보이스 번호 체크
            if (invoice_no == "" || invoice_no == null)
            {
                MessageBox.ShowMessage("인보이스번호를 입력해주세요.", this.Page);
                if (rbtnl_chk_process.SelectedValue == "view") //수정이면
                    tb_invoice_no.Focus();
                else
                    tb_new_invoice_no.Focus(); //신규생성시
            }
            else //인보이스를 저장한다.
            {
                sql = "select count(*) from SM_INVOICE_HDR_NEPES where invoice_no = '" + invoice_no + "' ";
                if (rbtnl_chk_process.SelectedValue == "view") //수정이면
                {

                    if (dbexe.QueryExecute(sql, "check") > 0) //기존 HDR에 저장된 값이 있다면 업데이트, 
                    {
                        sql = " update SM_INVOICE_HDR_NEPES set " +
                              "        invoice_type1     = '" + invoice_type1 + "' " +
                              "        , invoice_type2     = '" + invoice_type2 + "' " +
                              "        , ship_fr_cust_cd   = '" + ship_fr_cust_cd + "' " +
                              "        , ship_fr_cust_nm   = '" + ship_fr_cust_nm + "' " +
                              "        , ship_fr_add       = '" + ship_fr_add + "' " +
                              "        , ship_fr_tel       = '" + ship_fr_tel + "' " +
                              "        , ship_fr_fax       = '" + ship_fr_fax + "' " +
                              "        , bill_fr_cust_cd   = '" + bill_fr_cust_cd + "' " +
                              "        , bill_fr_cust_nm   = '" + bill_fr_cust_nm + "' " +
                              "        , bill_to_add       = '" + bill_to_add + "' " +
                              "        , bill_to_tel       = '" + bill_to_tel + "' " +
                              "        , bill_to_fax       = '" + bill_to_fax + "' " +
                              "        , bill_to_name      = '" + bill_to_name + "' " +
                              "        , ship_to_cust_cd   = '" + ship_to_cust_cd + "' " +
                              "        , ship_to_cust_nm   = '" + ship_to_cust_nm + "' " +
                              "        , ship_to_add       = '" + ship_to_add + "' " +
                              "        , ship_to_tel       = '" + ship_to_tel + "' " +
                              "        , ship_to_fax       = '" + ship_to_fax + "' " +
                              "        , ship_to_name      = '" + ship_to_name + "' " +
                              "        , port_of_loading   = '" + port_of_loading + "' " +
                              "        , final_destination = '" + final_destination + "' " +
                              "        , carrier           = '" + carrier + "' " +
                              "        , board_on_about    = '" + board_on_about + "' " +
                              "        , invoice_dt        = '" + invoice_dt + "' " +
                              "        , lc_no_dt          = '" + lc_no_dt + "' " +
                              "        , lc_issue_bank     = '" + lc_issue_bank + "' " +
                              "        , remark            = '" + remark + "' " +
                              "        , remark_incoterms  = '" + remark_incoterms + "' " +
                              "        , remark_pay_type   = '" + remark_pay_type + "' " +
                              "        , total_box_cnt     = '" + total_box_cnt + "' " +
                              "        , hts_code          = '" + hts_code + "' " +
                              "        , country_of_org    = '" + country_of_org + "' " +
                              "        , bank_cd           = '" + bank_cd + "' " +
                              "        , bank_name         = '" + bank_name + "' " +
                              "        , bank_addr         = '" + bank_addr + "' " +
                              "        , bank_branch       = '" + bank_branch + "' " +
                              "        , bank_swiftcode    = '" + bank_swiftcode + "' " +
                              "        , bank_acct_no      = '" + bank_acct_no + "' " +
                              "        , bank_accountee    = '" + bank_accountee + "' " +
                              "        , net_weight        = '" + net_weight + "' " +
                              "        , gross_weight      = '" + gross_weight + "' " +
                              "        , updt_user_id      = '" + Session["User"].ToString() + "' " +
                              "        , updt_dt           = getdate() " +
                              " where  invoice_no ='" + invoice_no + "' ";

                        if (dbexe.QueryExecute(sql, "") > 0)
                            MessageBox.ShowMessage("업데이트 되었습니다.", this.Page);
                    }
                    else
                    {
                        MessageBox.ShowMessage("저장된 Invoice가 없습니다. 생성버튼 클릭 후 진행해 주세요.", this.Page);
                    }
                }
                else
                {
                    if (dbexe.QueryExecute(sql, "check") > 0) //기존 HDR에 저장된 값이 없다면 , 
                    {
                        MessageBox.ShowMessage("이미 저장된 Invoice입니다. ", this.Page);
                        tb_new_invoice_no.Focus();
                    }
                    else
                    {
                        sql = " insert into SM_INVOICE_HDR_NEPES values ( " +
                       "  '" + invoice_no + "', '" + invoice_type1 + "', '" + invoice_type2 + "' " +
                       ", '" + ship_fr_cust_cd + "', '" + ship_fr_cust_nm + "', '" + ship_fr_add + "', '" + ship_fr_tel + "', '" + ship_fr_fax + "' " +
                       ", '" + bill_fr_cust_cd + "', '" + bill_fr_cust_nm + "', '" + bill_to_add + "', '" + bill_to_tel + "', '" + bill_to_fax + "' , '" + bill_to_name + "' " +
                       ", '" + ship_to_cust_cd + "', '" + ship_to_cust_nm + "', '" + ship_to_add + "', '" + ship_to_tel + "', '" + ship_to_fax + "' , '" + ship_to_name + "'" +
                       ", '" + port_of_loading + "', '" + final_destination + "', '" + carrier + "', '" + board_on_about + "', '" + invoice_dt + "' " +
                       ", '" + lc_no_dt + "', '" + lc_issue_bank + "', '" + remark + "', '" + remark_incoterms + "', '" + remark_pay_type + "', '" + total_box_cnt + "' " +
                       ", '" + hts_code + "', '" + country_of_org + "', '" + bank_cd + "', '" + bank_name + "', '" + bank_addr + "', '" + bank_branch + "', '" + bank_swiftcode + "', '" + bank_acct_no + "', '" + bank_accountee + "' " +
                       ", '" + net_weight + "' , '" + gross_weight + "' " +
                       ", '" + Session["User"].ToString() + "', getdate(), '" + Session["User"].ToString() + "', getdate()   )";

                        if (dbexe.QueryExecute(sql, "") > 0)
                            MessageBox.ShowMessage("저장되었습니다.", this.Page);
                    }
                }

            }
            btn_retrieve_Click(null, null);
        }

        protected void btn_retrieve_Click(object sender, EventArgs e)
        {
            string invoice_no, sql;
            invoice_no = tb_invoice_no.Text;
            // header 쪽 데이타 뿌려주기
            sql = " select invoice_no, invoice_type1, invoice_type2 " +
                  "      , ship_fr_cust_cd, ship_fr_cust_nm, ship_fr_add " +
                  "      , ship_fr_tel, ship_fr_fax, bill_fr_cust_cd " +
                  "      , bill_fr_cust_nm " +
                  "      , bill_to_add " +
                  "      , bill_to_tel " +
                  "      , bill_to_fax " +
                  "      , bill_to_name " +
                  "      , ship_to_cust_cd " +
                  "      , ship_to_cust_nm " +
                  "      , ship_to_add " +
                  "      , ship_to_tel " +
                  "      , ship_to_fax " +
                  "      , ship_to_name " +
                  "      , port_of_loading " +
                  "      , final_destination " +
                  "      , carrier " +
                  "      , board_on_about  " +
                  "      , invoice_dt " +
                  "      , lc_no_dt " +
                  "      , lc_issue_bank " +
                  "      , remark " +
                  "      , remark_incoterms " +
                  "      , remark_pay_type " +
                  "      , total_box_cnt " +
                  "      , hts_code " +
                  "      , country_of_org " +
                  "      , bank_cd " +
                  "      , bank_name   " +
                  "      , bank_addr    " +
                  "      , bank_branch   " +
                  "      , bank_swiftcode " +
                  "      , bank_acct_no   " +
                  "      , bank_accountee " +
                  "      , net_weight " +
                  "      , gross_weight " +
                  " from  SM_INVOICE_HDR_NEPES where invoice_no = '" + tb_invoice_no.Text + "' ";

            try
            {
                conn_erp.Open();
                cmd_erp = conn_erp.CreateCommand();
                cmd_erp.CommandType = CommandType.Text;
                cmd_erp.CommandText = sql;
                dr_erp = cmd_erp.ExecuteReader();
                if (dr_erp.HasRows == false)
                {
                    MessageBox.ShowMessage("조회된 내역이 없습니다.", this.Page);
                }
                else
                {
                    while (dr_erp.Read())
                    {
                        //tb_new_invoice_no.Text ;                              //인보이스번호
                        rbtnl_pay_type.SelectedValue = dr_erp[1].ToString();             //인보이스타입1(유상,무상,pra)
                        rbtnl_target_type.SelectedValue = dr_erp[2].ToString();           //인보이스타입2(고객,면허)
                        hf_tb_ship_fr_cust_cd.Value = dr_erp[3].ToString();  //발신인 거래처코드
                        tb_ship_fr_cust_nm.Text = dr_erp[4].ToString();
                        tb_ship_to_add.Text = dr_erp[5].ToString();                       //발신인주소
                        tb_ship_to_tel.Text = dr_erp[6].ToString();                        //발신인전화번호
                        tb_ship_fr_fax.Text = dr_erp[7].ToString();                        //발신인 fax
                        hf_tb_bill_to_cust_cd.Value = dr_erp[8].ToString();  //수취인 거래처코드
                        tb_bill_to_cust_nm.Text = dr_erp[9].ToString();
                        tb_bill_to_add.Text = dr_erp[10].ToString();                        //수취인 주소
                        tb_bill_to_tel.Text = dr_erp[11].ToString();                        //수취인 전화번호
                        tb_bill_to_fax.Text = dr_erp[12].ToString();                     //수취인 fax
                        tb_bill_to_name.Text = dr_erp[13].ToString();                     //수취인 
                        hf_tb_ship_to_cust_cd.Value = dr_erp[14].ToString();  //실물수령인 거래처코드
                        tb_ship_to_cust_nm.Text = dr_erp[15].ToString();
                        tb_ship_to_add.Text = dr_erp[16].ToString();                        //실물수령인 주소
                        tb_ship_to_tel.Text = dr_erp[17].ToString();                        //실물수령인 전화
                        tb_ship_to_fax.Text = dr_erp[18].ToString();                       //실물수령인 fax
                        tb_ship_to_name.Text = dr_erp[19].ToString();                       //실물수령인 fax
                        tb_port_of_loading.Text = dr_erp[20].ToString();                //출발지
                        tb_final_destination.Text = dr_erp[21].ToString();            //도착지
                        tb_carrier.Text = dr_erp[22].ToString();                               //운송업체
                        tb_board_on_about.Text = dr_erp[23].ToString();                  //발송일
                        tb_invoice_dt.Text = dr_erp[24].ToString();                           //인보이스 발행일
                        //null;                                           //신용장번호및발행일
                        //null;                                      //시용장개설은행
                        tb_remark.Text = dr_erp[27].ToString();                                   //비고
                        tb_remark_incoterms.Text = dr_erp[28].ToString();              //운임조건
                        tb_remark_pay_type.Text = dr_erp[29].ToString();               //유무상구분
                        tb_total_box_cnt.Text = dr_erp[30].ToString();                    //전체박스 수량
                        tb_hts_code.Text = dr_erp[31].ToString();                              //hts code
                        tb_country_of_org.Text = dr_erp[32].ToString();                  //원산지
                        tb_bank_info.Text = dr_erp[33].ToString();                             //회사 은행 코드
                        tb_bank_name.Text = dr_erp[34].ToString();                             //회사 은행명
                        tb_bank_addr.Text = dr_erp[35].ToString();                             //회사 은행 주소
                        tb_bank_branch.Text = dr_erp[36].ToString();                             //회사 은행 정보1
                        tb_bank_swiftcode.Text = dr_erp[37].ToString();                             //회사 은행 정보2
                        tb_bank_acct_no.Text = dr_erp[38].ToString();                             //회사 은행 정보3
                        tb_bank_accountee.Text = dr_erp[39].ToString();                             //회사 은행 정보4
                        tb_net_weight.Text = dr_erp[40].ToString();                         //net weight
                        tb_gross_weight.Text = dr_erp[41].ToString();                      //gross weight


                    }
                    dr_erp.Close();
                    conn_erp.Close();

                    Session["view"] = "OK";
                }
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
                Session["view"] = "NO";
            }
            conn_erp.Close();
            // detail 쪽 데이타 뿌려주기

            if (rbtnl_chk_process.SelectedValue == "view")
            {
                sql = "SELECT INVOICE_NO 인보이스번호 " +
                      "     , CUST_NM 거래처명 " +
                      "     , PL_DT PL날짜 " +
                      "     , PL_NO PL번호 " +
                      "     , PL_TYPE PL타입 " +
                      "     , ITEM_NM 디바이스명 " +
                      "     , CUST_LOT_NO 거래처LOT " +
                      "     , LOT_NO LOT번호 " +
                      "     , LOT_QTY 수량 " +
                      "     , LOT_UNIT 단위 " +
                      "     , PO_NO PO번호 " +
                      "     , FAB_LOT_NO  " +
                      "     , COST_TYPE COM_NONC " +
                      "     , COST_CUR_CD 단가화폐 " +
                      "     , COST_PRICE  단가 " +
                      "     , MAT_TYPE COM_NONC " +
                      "     , MAT_CUR_CD 원자재단가화폐 " +
                      "     , MAT_PRICE  원자재단가 " +
                      "     , ATTACH_TYPE COM_NONC " +
                      "     , ATTACH_CUR_CD 제3출하처용화폐 " +
                      "     , ATTACH_PRICE 제3출하처용단가 " +
                      "  FROM SM_INVOICE_DTL_NEPES " +
                      " WHERE invoice_no =  '" + tb_invoice_no.Text + "' ";
                try
                {                                
                    
                    erp_sqlAdapter = new SqlDataAdapter(sql, conn_erp);
                    ds = new DataSet();
                    erp_sqlAdapter.Fill(ds, "ds");
                    
                    FpSpread_view_data.DataSource = ds;
                    FpSpread_view_data.DataBind();
                    SetSpreadColumnLock_ViewData(11);
                }
                catch (Exception ex)
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                    //Session["view"] = "NO";
                }
            conn_erp.Close();

                ////if (conn_erp.State == ConnectionState.Open)
                ////    conn_erp.Close();
                ////erp_sqlAdapter = new SqlDataAdapter(sql, conn_erp);
                ////ds = new DataSet();
                ////erp_sqlAdapter.Fill(ds,"ds");

                //FpSpread_view_data.DataSource = ds;
                //FpSpread_view_data.DataBind();
                //SetSpreadColumnLock_ViewData(11); //불필요한부분 lock걸리게 하기
                ////SetSpreadDropDown_view(12); //단가화폐
                ////SetSpreadDropDown_view(14); //원자재화폐
                ////SetSpreadDropDown_view(16); //제3출하처화폐
            }
        }

        protected void cb_use_nepes_add_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_use_nepes_add.Checked)
            {
                string sql = "select BANK_CD, BANK_ENG_NM,  ENG_ADDR1 , ENG_ADDR2, ENG_ADDR3, '180-004-369451' BANK_ACCT,'NEPES CORPORATION' BIZ_NAME from B_BANK WHERE BANK_CD = '2600001'";
                // 은행정보 채우기
                try
                {
                    conn_erp.Open();
                    cmd_erp = conn_erp.CreateCommand();
                    cmd_erp.CommandType = CommandType.Text;
                    cmd_erp.CommandText = sql;
                    dr_erp = cmd_erp.ExecuteReader();

                    while (dr_erp.Read())
                    {
                        hf_tb_bank_cd.Value = dr_erp[0].ToString();
                        tb_bank_name.Text = dr_erp[1].ToString();
                        tb_bank_addr.Text = dr_erp[2].ToString();
                        tb_bank_branch.Text = dr_erp[3].ToString();
                        tb_bank_swiftcode.Text = dr_erp[4].ToString();
                        tb_bank_acct_no.Text = dr_erp[5].ToString();
                        tb_bank_accountee.Text = dr_erp[6].ToString();
                    }
                    dr_erp.Close();
                    conn_erp.Close();
                }
                catch { }
                finally
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                }
            }

        }

        protected void btn_spread_view_data_update_Click(object sender, EventArgs e)
        {
            FpSpread_view_data.SaveChanges();
        }

        protected void FpSpread_view_data_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            int colcnt;
            int r = (int)e.CommandArgument;
            string a_price_unit, b_price_unit, c_price_unit;
            string a_price, b_price, c_price;
            string a_type, b_type, c_type;
            string invoice_no, part, lot_number;
            colcnt = e.EditValues.Count - 1;

            //기준이 되는 값을 가져온다.
            invoice_no = tb_invoice_no.Text; //인보이스 번호
            part = FpSpread_view_data.Sheets[0].Cells[r, 5].Text; //조회된 품목명
            lot_number = FpSpread_view_data.Sheets[0].Cells[r, 7].Text; //조회된 lot번호

            // 수정할수 있는 row의 값을 가져온다.
            a_type = e.EditValues[12].ToString();
            a_price_unit = e.EditValues[13].ToString();
            a_price = e.EditValues[14].ToString();
            b_type = e.EditValues[15].ToString();
            b_price_unit = e.EditValues[16].ToString();
            b_price = e.EditValues[17].ToString();
            c_type = e.EditValues[18].ToString();
            c_price_unit = e.EditValues[19].ToString();
            c_price = e.EditValues[20].ToString();

            if (a_type == null || a_type == "" || a_type == "System.Object") //COST_CUR_CD
            {
                a_type = FpSpread_view_data.Sheets[0].Cells[r, 12].Text;
            }

            if (a_price_unit == null || a_price_unit == "" || a_price_unit == "System.Object") //COST_CUR_CD
            {
                a_price_unit = FpSpread_view_data.Sheets[0].Cells[r, 13].Text;
            }
            if (a_price == null || a_price == "" || a_price == "System.Object") //COST_PRICE
            {
                if (FpSpread_view_data.Sheets[0].Cells[r, 14] == null)
                   a_price = "0";
                else
                   a_price = FpSpread_view_data.Sheets[0].Cells[r, 14].Text;

            }
            if (b_type == null || b_type == "" || b_type == "System.Object") //COST_CUR_CD
            {
                b_type = FpSpread_view_data.Sheets[0].Cells[r, 15].Text;
            }
            if (b_price_unit == null || b_price_unit == "" || b_price_unit == "System.Object") //MAT_CUR_CD
            {

                b_price_unit = FpSpread_view_data.Sheets[0].Cells[r, 16].Text;
            }
            if (b_price == null || b_price == "" || b_price == "System.Object") //MAT_PRICE
            {
                if (FpSpread_view_data.Sheets[0].Cells[r, 17] == null)
                    b_price = "0";
                else
                    b_price = FpSpread_view_data.Sheets[0].Cells[r, 17].Text;
                
            }
            if (c_type == null || c_type == "" || c_type == "System.Object") //COST_CUR_CD
            {
                c_type = FpSpread_view_data.Sheets[0].Cells[r, 18].Text;
            }
            if (c_price_unit == null || c_price_unit == "" || c_price_unit == "System.Object") //APPROCH_CUR_CD
            {
                c_price_unit = FpSpread_view_data.Sheets[0].Cells[r, 19].Text;
            }
            if (c_price == null || c_price == "" || c_price == "System.Object") //APPROCH_PRICE
            {
                if (FpSpread_view_data.Sheets[0].Cells[r, 20] == null)
                    c_price = "0";
                else
                    c_price = FpSpread_view_data.Sheets[0].Cells[r, 20].Text;
                
            }

            string sql;
            sql = " update SM_INVOICE_DTL_NEPES " +
                  "    set cost_type = '" + a_type + "', MAT_TYPE = '" + b_type + "', ATTACH_TYPE = '" + c_type + "', cost_cur_cd = '" + a_price_unit + "', cost_price = '" + a_price + "' " +
                  "      , mat_cur_cd = '" + b_price_unit + "' , mat_price = '" + b_price + "' " +
                  "      , attach_cur_cd = '" + c_price_unit + "' , attach_price = '" + c_price + "' " +
                  "      , UPDT_DT = getdate() , updt_user_id = '" + Session["User"].ToString() + "' " +
                  "  where INVOICE_NO = '" + invoice_no + "' and item_nm = '" + part + "' and LOT_NO = '" + lot_number + "' ";

            if (dbexe.QueryExecute(sql, "") > 0)
                MessageBox.ShowMessage("업데이트 되었습니다.", this.Page);

            btn_retrieve_Click(null, null);

        }
        //네패스 체크박스 클릭시 해당 영문주소를 자동으로 발신인정보에 뿌려준다.
        protected void cb_nepes_addr_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_nepes_addr.Checked)
            {
                string sql;
                sql = "select BP_CD, BP_ENG_NM, ADDR1_ENG + ADDR2_ENG ADDR,TEL_NO1,FAX_NO from B_BIZ_PARTNER WHERE BP_CD = 'KO441'";

                // 정보 채우기
                try
                {
                    conn_erp.Open();
                    cmd_erp = conn_erp.CreateCommand();
                    cmd_erp.CommandType = CommandType.Text;
                    cmd_erp.CommandText = sql;
                    dr_erp = cmd_erp.ExecuteReader();

                    while (dr_erp.Read())
                    {
                        hf_tb_ship_fr_cust_cd.Value = dr_erp[0].ToString();
                        tb_ship_fr_cust_nm.Text = dr_erp[1].ToString();
                        tb_ship_fr_add.Text = dr_erp[2].ToString();
                        tb_ship_fr_tel.Text = dr_erp[3].ToString();
                        tb_ship_fr_fax.Text = dr_erp[4].ToString();
                    }
                    dr_erp.Close();
                    conn_erp.Close();
                }
                catch { }
                finally
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                }
            }
            else //초기화
            {
                hf_tb_ship_fr_cust_cd.Value = null;
                tb_ship_fr_cust_nm.Text = "";
                tb_ship_fr_add.Text = "";
                tb_ship_fr_tel.Text = "";
                tb_ship_fr_fax.Text = "";
            }

        }

        protected void btn_spread_view_data_delete_Click(object sender, EventArgs e)
        {
            //선택된 여러 row를 확인한다.
            System.Collections.IEnumerator enu = FpSpread_view_data.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;
            string invoice_no, part, lot_number;
            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread_view_data.Sheets[0].ActiveRow;
                //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                for (int i = 0; i < cr.RowCount; i++)
                {
                    //기준이 되는 값을 가져온다.
                    invoice_no = tb_invoice_no.Text; //인보이스 번호
                    part = FpSpread_view_data.Sheets[0].Cells[i, 5].Text; //조회된 품목명
                    lot_number = FpSpread_view_data.Sheets[0].Cells[i, 7].Text; //조회된 lot번호                   

                    string sql;
                    sql = "delete SM_INVOICE_DTL_NEPES " +
                          "  where INVOICE_NO = '" + invoice_no + "' and item_nm = '" + part + "' and LOT_NO = '" + lot_number + "' ";

                    if (dbexe.QueryExecute(sql, "") > 0)
                        MessageBox.ShowMessage("삭제 되었습니다.", this.Page);
                }
            }
            btn_retrieve_Click(null, null);
        }

        protected void btn_copy_Click(object sender, EventArgs e)
        {
            string sql, invoice_no, new_invoice_no;
            invoice_no = tb_invoice_no.Text;
            new_invoice_no = tb_new_invoice_no.Text;

            if (Session["view"].ToString() == "OK")
            {
                //신규인보이스번호 확인
                if (new_invoice_no == "" || new_invoice_no == null)
                {
                    MessageBox.ShowMessage("신규 인보이스 번호를 입력하세요..", this.Page);
                    tb_new_invoice_no.Focus();
                }
                else if (tb_invoice_no.Text == "" || tb_invoice_no.Text == null)
                {
                    MessageBox.ShowMessage("원본 인보이스 번호를 입력하세요..", this.Page);
                    tb_invoice_no.Focus();
                }
                else
                {
                    sql = "insert into SM_INVOICE_HDR_NEPES " +
                        "  select '" + new_invoice_no + "'" +
                        "       , invoice_type1,invoice_type2,ship_fr_cust_cd,ship_fr_cust_nm,ship_fr_add,ship_fr_tel,ship_fr_fax " +
                        "       , bill_fr_cust_cd,bill_fr_cust_nm,bill_to_add,bill_to_tel,bill_to_fax,bill_to_name " +
                        "       , ship_to_cust_cd,ship_to_cust_nm,ship_to_add,ship_to_tel,ship_to_fax,ship_to_name " +
                        "       , port_of_loading,final_destination,carrier,board_on_about,invoice_dt,lc_no_dt " +
                        "       , lc_issue_bank,remark,remark_incoterms,remark_pay_type,total_box_cnt,hts_code " +
                        "       , country_of_org,bank_cd,bank_name,bank_addr,bank_branch,bank_swiftcode,bank_acct_no,bank_accountee " +
                        "       , net_weight,gross_weight, '" + Session["User"].ToString() + "', getdate(), '" + Session["User"].ToString() + "', getdate()  " +
                        "   from SM_INVOICE_HDR_NEPES where invoice_no = '" + tb_invoice_no.Text + "' ";

                    if (dbexe.QueryExecute(sql, "") > 0)
                        MessageBox.ShowMessage("저장 되었습니다.", this.Page);
                }
            }
            else
            {
                MessageBox.ShowMessage("Copy하려는 인보이스 번호를 입력 후 조회해 주세요.", this.Page);
            }

        }

        protected void btn_preview_Click(object sender, EventArgs e)
        {
            // 단가 체크 개수를 파악한다. 
            int chk_cnt = 0;
            if (cb_price1.Checked)
                chk_cnt = chk_cnt + 1; //단가
            if (cb_price2.Checked)
                chk_cnt = chk_cnt + 1; //원자재
            if (cb_price3.Checked)
                chk_cnt = chk_cnt + 1; //3자국
            if (chk_cnt > 0 && chk_cnt < 3)
            {
                string sql =
                   "  select invoice_no " +
                   "       , invoice_type1,invoice_type2,ship_fr_cust_cd,ship_fr_cust_nm,ship_fr_add,ship_fr_tel,ship_fr_fax " +
                   "       , bill_fr_cust_cd,bill_fr_cust_nm,bill_to_add,bill_to_tel,bill_to_fax,bill_to_name " +
                   "       , ship_to_cust_cd,ship_to_cust_nm,ship_to_add,ship_to_tel,ship_to_fax,ship_to_name " +
                   "       , port_of_loading,final_destination,carrier,board_on_about,invoice_dt,lc_no_dt " +
                   "       , lc_issue_bank,remark,remark_incoterms,remark_pay_type,total_box_cnt,hts_code " +
                   "       , country_of_org,bank_cd,bank_name,bank_addr,bank_branch,bank_swiftcode,bank_acct_no,bank_accountee " +
                   "       , net_weight,gross_weight " +
                   "   from SM_INVOICE_HDR_NEPES where invoice_no = '" + tb_invoice_no.Text + "' ";

                ds_sm_s3001 dt1 = new ds_sm_s3001();
                ReportViewer1.Reset();


                //상세내용을 뿌려준다.
                //ReportCreator("DTL", dt1, "dbo.USP_SM_S3001_MES_INVOICE_DTL_VIEW", ReportViewer1, "rv_sm_s3001_no1_dtl.rdlc", "DataSet1");
                // header 를 뿌려준다.
                ReportCreator("HDR", dt1, sql, ReportViewer1, "rv_sm_s3001_no1.rdlc", "DataSet1");

            }
            else
            {
                MessageBox.ShowMessage("단가는 1~2개로 선택하여 주세요.. ", this.Page);
            }

            //{
            //    string sub_sql;

            //    sub_sql = "SELECT COST_TYPE,COST_CUR_CD,COST_PRICE" +
            //              " FROM SM_INVOICE_DTL_NEPES" +
            //              " WHERE COST_TYPE IS NOT NULL ";

            //    if (chb_price_view_type.SelectedValue.ToString() == "A")
            //    {
            //        string sql;
            //        sql = "SELECT PL_TYPE,ITEM_NM,LOT_NO,SUM(LOT_QTY),LOT_UNIT" +
            //              " FROM SM_INVOICE_DTL_NEPES" +
            //              " GROUP BY PL_TYPE, ITEM_NM, LOT_NO,LOT_UNIT";
            //    }

            //}
        }




        protected void btn_select_all_Click(object sender, EventArgs e)
        {
            
            ////조회된 내역이 있어야만 한다.
            //if (FpSpread_new_data.Sheets[0].Rows.Count > 0)
            //{
            //    //전체선택시 체크버튼
            //    if (btn_select_all.Text == "전체선택")
            //    {
            //        for (int i = 0; i <= FpSpread_new_data.Rows.Count - 1; i++)
            //        {
                        
            //            //FpSpread_new_data.Sheets[0].Cells[i, 0].CanFocus = true;
            //            //FpSpread_new_data.Sheets[0].Cells[i, 0].ResetCanFocus();
            //            FarPoint.Web.Spread.Model.ISheetSelectionModel model = FpSpread_new_data.ActiveSheetView.SelectionModel;
            //            model.AddSelection(i, 0, 1, 1);
            //            bool chk;
            //            chk = model.IsSelected(i,1);
            //            if (model.IsSelected(i, 0))
            //                FpSpread_new_data.Sheets[0].Cells[model.AnchorRow, 0].Value = "1"; //체크박스에 체크

            //           // FpSpread_new_data.ActiveSheetView.SelectionModel.SetSelection(i, 0, 1, 1);
            //           // System.Collections.IEnumerator enu = FpSpread_new_data.ActiveSheetView.SelectionModel.GetEnumerator();
            //           // FarPoint.Web.Spread.Model.CellRange cr;
            //           //// FarPoint.Web.Spread.Model.h
            //           // while (enu.MoveNext())
            //           // {

            //           //     cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
            //           //     //FpSpread_new_data.Sheets[0].Cells[cr.Row, 0].Value = "1"; //체크박스에 체크
            //           //     FpSpread_new_data.Sheets[0].Cells[cr.Row, 0].Text = "True"; //.Value = "1"; //체크박스에 체크
            //           //     FpSpread_new_data.EditModePermanent = true;
            //           //     //FpSpread_new_data.CellClick
            //           // }

            //           // FpSpread_new_data.se
            //            //FpSpread_new_data.Sheets[0].Cells[i,0]
            //        }
            //        btn_select_all.Text = "전체해제";
            //    }
            //    else if (btn_select_all.Text == "전체해제")
            //    {
            //        for (int i = 0; i <= FpSpread_new_data.Rows.Count - 1; i++)
            //        {
            //            FarPoint.Web.Spread.Model.ISheetSelectionModel model = FpSpread_new_data.Sheets[0].SelectionModel;
            //            model.RemoveSelection(i, 0, 1, 1);
            //            //FpSpread_new_data.Sheets[0].Cells[i, 0].Value = "0"; //체크박스에 체크해제
            //            //FpSpread_new_data.EditModePermanent = false;

            //        }
            //        btn_select_all.Text = "전체선택";
            //    }
            //    UpdatePanel_body2.Update();
            //}
            //else
            //{
            //    MessageBox.ShowMessage("조회된 데이타가 없습니다. ", this.Page);
            //}
            
        }

        private void ReportCreator(string gubun, DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {
            DataSet ds = _dataSet;
            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();


            // header 내용을 보여준다.
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = _Query;
            dr_erp = cmd_erp.ExecuteReader();
            ds.Tables[0].Load(dr_erp);


            try
            {
                dr_erp.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                _reportViewer.LocalReport.DataSources.Add(rds);
                String view_type1 = "NULL",view_type2 = "NULL";
                if (cb_price1.Checked)
                    view_type1 = "A"; //단가
                if (cb_price2.Checked)
                    view_type2 = "B"; //원자재
                if (cb_price3.Checked)
                    view_type2 = "C"; //3자국

                _reportViewer.LocalReport.SetParameters(new ReportParameter("view_type1", view_type1));
                _reportViewer.LocalReport.SetParameters(new ReportParameter("view_type2", view_type2));
                _reportViewer.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
                _reportViewer.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing_pono);


                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }

        }

        // sub 레포트 데이타 연결
        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();

            //dtl 내용을 sp에서 가져다가 보여준다.
            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.USP_SM_S3001_MES_INVOICE_DTL_VIEW2";
            cmd_erp.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@INVOICE_NO", SqlDbType.VarChar, 30);
            SqlParameter param2 = new SqlParameter("@VIEW_TYPE", SqlDbType.VarChar, 10);
            SqlParameter param3 = new SqlParameter("@SELECT_TYPE", SqlDbType.VarChar, 10);
            param1.Value = tb_invoice_no.Text;
            // 상세보기(LOT출력), 집계보기 체크한다. 
            if (cb_view_lot.Checked)
                param2.Value = "DTL";
            else
                param2.Value = "HDR";

            //단가선택유형을 체크한다.
            string param3_data = "";
            if (cb_price1.Checked)
                param3_data = param3_data + "A"; //단가
            if (cb_price2.Checked)
                param3_data = param3_data + "B"; //원자재
            if (cb_price3.Checked)
                param3_data = param3_data + "C"; //3자국

            param3.Value = param3_data;

            cmd_erp.Parameters.Add(param1);
            cmd_erp.Parameters.Add(param2);
            cmd_erp.Parameters.Add(param3);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                DataTable dt = new DataTable();
                da.Fill(dt);
                //string report_nm;
                //report_nm = "rv_sm_s3001_no1_sub.rdlc";
                //ReportViewer1.LocalReport.ReportPath = Server.MapPath(report_nm);
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                //변수 전달
                //_reportViewer.LocalReport.SetParameters(new ReportParameter("ld_in_qty", ld_in_qty.ToString()));

                //ReportViewer1.LocalReport.DataSources.Add(rds);
                //ReportViewer1.LocalReport.Refresh();
                e.DataSources.Add(new ReportDataSource("DataSet1", dt));
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
        }

        // sub 레포트 데이타 연결
        void LocalReport_SubreportProcessing_pono(object sender, SubreportProcessingEventArgs e)
        {

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();

            //dtl 내용을 sp에서 가져다가 보여준다.
            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.USP_SM_INVOICE_DTL_PO_NO";
            cmd_erp.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@INVOICE_NO", SqlDbType.VarChar, 30);

            if (cb_view_pono.Checked) //po보기 체크시 실행
                param1.Value = tb_invoice_no.Text;
            else
                param1.Value = "";
            cmd_erp.Parameters.Add(param1);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                DataTable dt = new DataTable();
                //ds_sm_s3001_no2_sub ds = new ds_sm_s3001_no2_sub();

                da.Fill(dt);
                
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "ds_sm_s3001_no2_sub";
                rds.Value = dt;

                e.DataSources.Add(new ReportDataSource("ds_sm_s3001_no2_sub", dt));
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
        }


        protected void FpSpread_new_data_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            //int row = (int)e.CommandArgument;
            //string a;
            //a = FpSpread_new_data.Sheets[0].ColumnHeader.Cells[row, 0].Text;
            ////if (FpSpread_new_data.Sheets[0]..Cells[r, 0]..Value = "1"; //체크박스에 체크
        }

        //private void Sub_ReportCreator(string gubun, DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        //{
        //    DataSet ds = _dataSet;
        //    conn_erp.Open();
        //    cmd_erp = conn_erp.CreateCommand();

        //    //dtl 내용을 sp에서 가져다가 보여준다.
        //    cmd_erp.CommandType = CommandType.StoredProcedure;
        //    cmd_erp.CommandText = _Query;
        //    cmd_erp.CommandTimeout = 3000;
        //    SqlParameter param1 = new SqlParameter("@INVOICE_NO", SqlDbType.VarChar, 30);
        //    SqlParameter param2 = new SqlParameter("@VIEW_TYPE", SqlDbType.VarChar, 10);
        //    param1.Value = tb_invoice_no.Text;
        //    if (cb_view_lot.Checked)
        //        param2.Value = "DTL";
        //    else
        //        param2.Value = "HDR";

        //    cmd_erp.Parameters.Add(param1);
        //    cmd_erp.Parameters.Add(param2);

        //    erp_sqlAdapter = new SqlDataAdapter(cmd_erp);
        //    DataTable dt = new DataTable();
        //    erp_sqlAdapter.Fill(ds.Tables[0]);

        //    try
        //    {
        //        //cmd_erp.CommandText = _Query;
        //        //dr_erp = cmd_erp.ExecuteReader();
        //        //ds.Tables[0].Load(dr_erp);
        //        dr_erp.Close();
        //        _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

        //        _reportViewer.LocalReport.DisplayName = "REPORT_" + DateTime.Now.ToShortDateString();
        //        ReportDataSource rds = new ReportDataSource();
        //        rds.Name = _ReportDataSourceName;
        //        rds.Value = ds.Tables[0];
        //        _reportViewer.LocalReport.DataSources.Add(rds);
        //        _reportViewer.LocalReport.Refresh();
        //    }
        //    catch { }
        //    finally
        //    {
        //        if (conn_erp.State == ConnectionState.Open)
        //            conn_erp.Close();
        //    }

        //}

    }
    //전체선택용 체크박스
    [Serializable()]
    public class myCheck : FarPoint.Web.Spread.CheckBoxCellType
    {
        public override System.Web.UI.Control PaintCell(string id, System.Web.UI.WebControls.TableCell parent, FarPoint.Web.Spread.Appearance style, FarPoint.Web.Spread.Inset margin, object value, bool upperLevel)
        {
            Control c = base.PaintCell(id, parent, style, margin, value, upperLevel);
            CheckBox chk = (CheckBox)c.Controls[0];
            chk.Attributes.Add("onclick", "myCheckFunction()");
            return c;
        }
    }
}