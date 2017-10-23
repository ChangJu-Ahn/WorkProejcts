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
using System.IO;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition;
using System.Drawing;
using ERPAppAddition.ERPAddition.SM.sm_sa001;

namespace ERPAppAddition.ERPAddition.SM.sm_sa001
{
  
    public partial class sm_sa001 : System.Web.UI.Page
    {
        int value;

        sa_fun fun = new sa_fun();

        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();        
        string lot_no;
        bool insertFlag = false;

        FarPoint.Web.Spread.CheckBoxCellType chkboxCelltype = new FarPoint.Web.Spread.CheckBoxCellType();


        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;                
                
                setCombo();

                TXT_INPUTDT.Text = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
                DateTime setDate = DateTime.Today.AddDays(-7);
                tb_fr_yyyymmdd.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00") + setDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
                grid();

                WebSiteCount();
            }
            
        }


        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void grid()
        {

            FpSpread_new_data.ActiveSheetView.Reset();
            FpSpread_new_data.ActiveSheetView.ColumnCount = 29;
            FpSpread_new_data.ActiveSheetView.Rows.Count = 0;
            FpSpread_new_data.ActiveSheetView.AllowPage = false;
            FarPoint.Web.Spread.DefaultSkins.Classic.Apply(FpSpread_new_data);
            
            FpSpread_new_data.Columns[0].Width = 35;
            FpSpread_new_data.Columns[1].Width = 35;
            FpSpread_new_data.Columns[2].Width = 200;
            FpSpread_new_data.Columns[3].Width = 100;
            FpSpread_new_data.Columns[4].Width = 70;
            FpSpread_new_data.Columns[5].Width = 120;
            FpSpread_new_data.Columns[6].Width = 70;
            FpSpread_new_data.Columns[7].Width = 90;
            FpSpread_new_data.Columns[8].Width = 60;
            FpSpread_new_data.Columns[9].Width = 120;
            FpSpread_new_data.Columns[10].Width = 60;
            FpSpread_new_data.Columns[11].Width = 60;
            FpSpread_new_data.Columns[12].Width = 45;
            FpSpread_new_data.Columns[13].Width = 150;
            FpSpread_new_data.Columns[14].Width = 40;
            FpSpread_new_data.Columns[15].Width = 40;
            FpSpread_new_data.Columns[16].Width = 40;
            FpSpread_new_data.Columns[17].Width = 95;
            FpSpread_new_data.Columns[18].Width = 70;
            FpSpread_new_data.Columns[19].Width = 60;
            FpSpread_new_data.Columns[20].Width = 120;
            FpSpread_new_data.Columns[21].Width = 60;
            FpSpread_new_data.Columns[22].Width = 60;
            FpSpread_new_data.Columns[23].Width = 38;
            FpSpread_new_data.Columns[24].Width = 120;
            FpSpread_new_data.Columns[25].Width = 200;
            FpSpread_new_data.Columns[26].Width = 60;
            FpSpread_new_data.Columns[27].Width = 40;
            FpSpread_new_data.Columns[28].Width = 250;

            FpSpread_new_data.Columns[0].DataField = "CHK";
            FpSpread_new_data.Columns[1].DataField = "C_CHK";
            FpSpread_new_data.Columns[2].DataField = "LOT_NO";
            FpSpread_new_data.Columns[3].DataField = "DRAIN_IN_DT";
            FpSpread_new_data.Columns[4].DataField = "DRAIN_PLANT";
            FpSpread_new_data.Columns[5].DataField = "DRAIN_PLANT_NM";
            FpSpread_new_data.Columns[6].DataField = "DRAIN_PROCESS";
            FpSpread_new_data.Columns[7].DataField = "DRAIN_MACHINE";
            FpSpread_new_data.Columns[8].DataField = "DRAIN_MAT";
            FpSpread_new_data.Columns[9].DataField = "DRAIN_MAT_NM";
            FpSpread_new_data.Columns[10].DataField = "DRAIN_QTY";
            FpSpread_new_data.Columns[11].DataField = "DRAIN_SCRAP_QTY";
            FpSpread_new_data.Columns[12].DataField = "DRAIN_UINT";
            FpSpread_new_data.Columns[13].DataField = "DRAIN_RMK";
            FpSpread_new_data.Columns[14].DataField = "OUT_YN";
            FpSpread_new_data.Columns[15].DataField = "SEQ";
            FpSpread_new_data.Columns[16].DataField = "STATE_FLAG";
            FpSpread_new_data.Columns[17].DataField = "R_DT";
            FpSpread_new_data.Columns[18].DataField = "R_INPUT_DT";
            FpSpread_new_data.Columns[19].DataField = "R_MAT";
            FpSpread_new_data.Columns[20].DataField = "R_MAT_NM";
            FpSpread_new_data.Columns[21].DataField = "R_QTY";
            FpSpread_new_data.Columns[22].DataField = "R_QTY_UNIT";
            FpSpread_new_data.Columns[23].DataField = "R_CUST";
            FpSpread_new_data.Columns[24].DataField = "R_CUST";
            FpSpread_new_data.Columns[25].DataField = "R_DOC_NO";
            FpSpread_new_data.Columns[26].DataField = "IN_AU_QTY";
            FpSpread_new_data.Columns[27].DataField = "IN_AU_UNIT";
            FpSpread_new_data.Columns[28].DataField = "R_RMK";                    
           
            FpSpread_new_data.Columns[0].Label = "확인";
            FpSpread_new_data.Columns[1].Label = "취소확인";
            FpSpread_new_data.Columns[2].Label = "LOT No.";
            FpSpread_new_data.Columns[3].Label = "Scrap발생일";
            FpSpread_new_data.Columns[4].Label = "발생공장CD";
            FpSpread_new_data.Columns[4].Visible = false;
            FpSpread_new_data.Columns[5].Label = "발생공장";
            FpSpread_new_data.Columns[6].Label = "발생공정";
            FpSpread_new_data.Columns[7].Label = "발생장비";
            FpSpread_new_data.Columns[8].Label = "Scrap종류CD";
            FpSpread_new_data.Columns[8].Visible = false;
            FpSpread_new_data.Columns[9].Label = "Scrap종류";
            FpSpread_new_data.Columns[10].Label = "누적장수";
            FpSpread_new_data.Columns[11].Label = "Scrap수량";
            FpSpread_new_data.Columns[12].Label = "단위";
            FpSpread_new_data.Columns[13].Label = "비고";
            FpSpread_new_data.Columns[14].Label = "반출유무";
            FpSpread_new_data.Columns[14].Visible = false;
            FpSpread_new_data.Columns[15].Label = "SEQ";
            FpSpread_new_data.Columns[15].Visible = false;
            FpSpread_new_data.Columns[16].Label = "FLAG";
            FpSpread_new_data.Columns[16].Visible = false;
            FpSpread_new_data.Columns[17].Label = "반출일";
            FpSpread_new_data.Columns[18].Label = "입력일";
            FpSpread_new_data.Columns[19].Label = "반출품목CD";
            FpSpread_new_data.Columns[19].Visible = false;
            FpSpread_new_data.Columns[20].Label = "반출품목";
            FpSpread_new_data.Columns[21].Label = "반출수량";
            FpSpread_new_data.Columns[22].Label = "단위";
            FpSpread_new_data.Columns[23].Label = "반출처CD";
            FpSpread_new_data.Columns[23].Visible = false;
            FpSpread_new_data.Columns[24].Label = "반출처";
            FpSpread_new_data.Columns[25].Label = "Scrap반출번호";
            FpSpread_new_data.Columns[26].Label = "Au농도";
            FpSpread_new_data.Columns[27].Label = "단위";
            FpSpread_new_data.Columns[28].Label = "입력 비고";

            //FpSpread_new_data.ActiveSheetView.Columns[0].CellType = chkboxCelltype;
            //FpSpread_new_data.ActiveSheetView.Columns[1].CellType = chkboxCelltype; 
            for (int c = 0; c < FpSpread_new_data.Columns.Count; c++)
            {
                SetSpreadColumnLock(c);
            }            
        }

        private void setCombo()
        {
            /*au농도 단위*/
            string[] AUUNIT = { "", "g/KG", "g/LT" };
            DDL_AUUNIT.DataSource = AUUNIT;
            DDL_AUUNIT.DataBind();

            /*공장*/
            DataTable plant = fun.getData("select DISTINCT PLANT_CD, PLANT_DESC from dbo.SA_SYS_CODE");
            if (plant.Rows.Count > 0)
            {
                SDDL_FAC.DataTextField = "PLANT_DESC";
                SDDL_FAC.DataValueField = "PLANT_CD";
                SDDL_FAC.DataSource = plant;
                SDDL_FAC.DataBind();
                SDDL_FAC.SelectedIndex = 0;
            }

            seqComBoSet("FAC");
            DDL_MAT.SelectedIndex = 0;
            DDL_MACH.SelectedIndex = 0;

            /*수량단위 */
            DataTable UNIT = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8005'");
            if (UNIT.Rows.Count > 0)
            {
                DataRow dr = UNIT.NewRow();
                UNIT.Rows.InsertAt(dr, 0);

                DDL_UNIT.DataTextField = "MINOR_NM";
                DDL_UNIT.DataValueField = "MINOR_CD";
                DDL_UNIT.DataSource = UNIT;
                DDL_UNIT.DataBind();
            }

            /*반출처 */            
            DataTable CUST = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004'");
            if (CUST.Rows.Count > 0)
            {
                DataTable ddt = CUST.Copy();
                DataRow dr = CUST.NewRow();
                CUST.Rows.InsertAt(dr, 0);

                DDL_CUST.DataTextField = "MINOR_NM";
                DDL_CUST.DataValueField = "MINOR_CD";
                DDL_CUST.DataSource = CUST;
                DDL_CUST.DataBind();
            }
        }

        protected void search()
        {
            grid();
            DataSet ds = new DataSet();
            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    DataRow dr = ds.Tables["DataSet1"].NewRow();
                    ds.Tables["DataSet1"].Rows.InsertAt(dr, 0);
                    FpSpread_new_data.DataSource = ds.Tables["DataSet1"];
                    FpSpread_new_data.DataBind();

                    txt_all_sum.Text = "";
                    txt_sum.Text = "";
                    txt_c_sum.Text = "";

                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                
                FpSpread_new_data.DataSource = ds.Tables["DataSet1"];
                FpSpread_new_data.DataBind();
                DataTable dts = ds.Tables["DataSet1"].Copy();

                double all_sum = 0;

                for(int j = 0; j < dts.Rows.Count; j++)
                {
                    FpSpread_new_data.ActiveSheetView.Cells[j, 0].CellType = chkboxCelltype;
                    FpSpread_new_data.ActiveSheetView.Cells[j, 1].CellType = chkboxCelltype;

                    /*전체 조회 체크박스 선택시 chk못하게 lock*/
                    if(CheckBox1.Checked == true)
                    {
                        FpSpread_new_data.Cells[j, 0].Locked = true;
                        FpSpread_new_data.Cells[j, 1].Locked = true;                        
                    }
                    else
                    {
                        if (dts.Rows[j]["OUT_YN"].ToString() == "Y")
                        {
                            FpSpread_new_data.Cells[j, 0].Locked = true;
                            FpSpread_new_data.Cells[j, 1].Locked = false;                            
                        }
                        else
                        {
                            FpSpread_new_data.Cells[j, 0].Locked = false;
                            FpSpread_new_data.Cells[j, 1].Locked = true;                           
                        }
                    }

                    double val = Convert.ToDouble(ds.Tables[0].Rows[j]["DRAIN_SCRAP_QTY"].ToString());
                    all_sum = all_sum + val;
                }

                txt_all_sum.Text = all_sum.ToString("0.00");

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }

        }
        private void SetSpreadColumnLock(int column)
        {
            if (column != 2)
            {                
                FpSpread_new_data.ActiveSheetView.Columns[column].HorizontalAlign = HorizontalAlign.Center;                
            }
            if (column >= 2) 
            {
                FpSpread_new_data.ActiveSheetView.Columns[column].Locked = true;
            }
            
            FpSpread_new_data.ActiveSheetView.Columns[column].VerticalAlign = VerticalAlign.Middle;
            FpSpread_new_data.ActiveSheetView.Columns[column].Font.Name = "돋움체";
            FpSpread_new_data.ActiveSheetView.LockBackColor = Color.LightCyan;
            FpSpread_new_data.ActiveSheetView.LockForeColor = Color.Black;
            
        }
        private string getSQL()
        {  
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select             				  \n");
            sbSQL.Append(" 0 AS CHK                     \n");
            sbSQL.Append(",0 AS C_CHK                   \n");
            sbSQL.Append(",DD.LOT_NO                          \n");            
            sbSQL.Append(",DD.DRAIN_IN_DT			          \n");
            sbSQL.Append(",DD.DRAIN_PLANT                     \n");
            sbSQL.Append(",(select TOP 1 PLANT_DESC from SA_SYS_CODE WHERE PLANT_CD = DD.DRAIN_PLANT) AS DRAIN_PLANT_NM   /*20151229 공장 명으로 변경  기존 : DRAIN_PLANT*/ \n");
            sbSQL.Append(",DD.DRAIN_PROCESS                   \n");
            sbSQL.Append(",DD.DRAIN_MACHINE                   \n");
            sbSQL.Append(",DD.DRAIN_MAT                       \n");
            sbSQL.Append(",(SELECT top 1 MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8003' and MINOR_CD = DD.DRAIN_MAT) AS DRAIN_MAT_NM /* 20151229 코드명으로 출력 이근만 SCRAP종류 */ \n");
            sbSQL.Append(",DD.DRAIN_QTY                       \n");
            sbSQL.Append(",DD.DRAIN_SCRAP_QTY                 \n");
            sbSQL.Append(",DD.DRAIN_UINT                      \n");
            sbSQL.Append(",DD.DRAIN_RMK                       \n");
            sbSQL.Append(",DD.OUT_YN                          \n");
            sbSQL.Append(",DD.SEQ                             \n");
            sbSQL.Append(",HH.STATE_FLAG			          \n");
            sbSQL.Append(",HH.R_DT                            \n");
            sbSQL.Append(",HH.R_INPUT_DT                      \n");
            sbSQL.Append(",HH.R_MAT                           \n");
            sbSQL.Append(",(select TOP 1 GROUP1_DESC from SA_SYS_CODE WHERE PLANT_CD = DD.DRAIN_PLANT AND HH.R_MAT = GROUP1_CODE) AS R_MAT_NM --20151221 MAT명으로 수정 HH.R_MAT \n");
            sbSQL.Append(",HH.R_QTY                           \n");
            sbSQL.Append(",HH.R_QTY_UNIT                      \n");
            sbSQL.Append(",HH.R_CUST                          \n");
            sbSQL.Append(",(SELECT TOP 1 MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004' AND MINOR_CD = HH.R_CUST) AS R_CUST_NM /*20151202 코드명으로 출력 이근만 업체명*/ \n");
            sbSQL.Append(",HH.R_DOC_NO                        \n");
            sbSQL.Append(",HH.IN_AU_QTY                       \n");
            sbSQL.Append(",HH.IN_AU_UNIT                      \n");
            sbSQL.Append(",HH.R_RMK                           \n");            
            sbSQL.Append("from OUT_MAT_DRAIN DD               \n");
            sbSQL.Append("    ,OUT_MAT_HIS HH                 \n");
            sbSQL.Append("WHERE DD.LOT_NO *= HH.LOT_NO        \n");
            sbSQL.Append("  AND SUBSTRING(DD.DRAIN_IN_DT, 1, 8) BETWEEN '" + tb_fr_yyyymmdd.Text + "' AND '" + tb_to_yyyymmdd.Text + "' \n");
            sbSQL.Append("  AND( DD.COMF1_DT IS NOT NULL AND DD.COMF1_DT != '')       \n");  //확인란에 DATE가 있을때만 LOT생성 대상이된다.
            sbSQL.Append("  AND( DD.COMF2_DT IS NOT NULL AND DD.COMF2_DT != '')       \n");  //확인란에 DATE가 있을때만 LOT생성 대상이된다.            
            sbSQL.Append("  AND DD.STATE_FLAG <> 'D'  \n");  //확인란에 DATE가 있을때만 LOT생성 대상이된다.           

            if(CheckBox1.Checked == false)
            {
                sbSQL.Append("  AND DD.DRAIN_PLANT = '" + SDDL_FAC .Text + "'        \n");
                sbSQL.Append("  AND DD.DRAIN_MACHINE = '" + DDL_MACH.Text + "'        \n");
                sbSQL.Append("  AND DD.DRAIN_MAT = '" + DDL_MAT.Text + "'        \n");
            }
            if (CheckBox2.Checked == true)
            {
                sbSQL.Append("  AND DD.LOT_NO IS NULL \n");                
            }

            return sbSQL.ToString();
        }

        public void initTextBox()
        {            
            TXT_RDT.Text = "";            
            DDL_CUST.Text = "";
            TXT_SCRQTY.Text = "";
            TXT_DOCNO.Text = "";
            TXT_RMK.Text = "";
            DDL_UNIT.Text = "";
            TXT_AUQTY.Text = "";
            DDL_AUUNIT.Text = "";
        }

       

        protected void btn_mighty_retrieve_Click1(object sender, EventArgs e)
        {            
            search();            
            initTextBox();            
        }

        protected void FpSpread_new_data_UpdateCommand(DataTable gridData)
        {
            string r_dt = TXT_RDT.Text;
            string r_inputDt = TXT_INPUTDT.Text;
            string r_mat = DDL_MAT.Text;
            string r_cust = DDL_CUST.Text;
            string r_qty = TXT_SCRQTY.Text;
            string r_unit = DDL_UNIT.Text;
            string r_docNo = TXT_DOCNO.Text;
            string AUQTY = TXT_AUQTY.Text == "" ? "0" : TXT_AUQTY.Text;
            string AUUNIT = DDL_AUUNIT.Text;
            string r_rmk = TXT_RMK.Text;

            if (lot_no == "" || lot_no == null)
            {
                lot_no = r_inputDt + r_cust + SDDL_FAC.Text + DDL_MACH.Text + DDL_MAT.Text;
                //입력일 / 반출처 / 공장 / 발생장비 / scrap품목

                /*lot번호 생성후 table 뒤져서 lot번호 체번*/
                string setSQL = "select max(lot_no) as OUT from OUT_MAT_HIS where LOT_NO LIKE '" + lot_no + "%'";
                DataSet ds = new DataSet();

                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = setSQL;

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                string Out = ds.Tables["DataSet1"].Rows[0]["OUT"].ToString() == null ? "" : ds.Tables["DataSet1"].Rows[0]["OUT"].ToString();
                /*값이 나온다면끝자리에  + 1  예)[-02] 붙이기 */
                if (Out != "")
                {
                    string aa = Out.Substring(lot_no.Length, Out.Length - lot_no.Length);
                    if (aa.Length == 0)
                    {
                        lot_no = lot_no + "-02";
                    }
                    else
                    {
                        string add = aa.Replace("-", "");
                        lot_no = lot_no + "-" + (Convert.ToInt32(add) + 1).ToString("00");
                    }
                }

            }

            /*INSERT해야할 DELETE 해야할 데이터 ROWS를 찾는다.*/
            DataTable dtInsert0 = gridData.Copy();
            DataTable dtDelete0 = gridData.Copy();
            dtInsert0.DefaultView.RowFilter = "CHK = '1'";
            DataTable dtInsert = dtInsert0.DefaultView.ToTable();
            dtDelete0.DefaultView.RowFilter = "C_CHK = '1'";
            DataTable dtDelete = dtDelete0.DefaultView.ToTable();


            for (int r = 0; r < dtInsert.Rows.Count; r++)
            {
                //저장 체크된 로직인지 확인하기
                if (dtInsert.Rows[r]["CHK"].ToString() == "1") //체크된 row이면
                {
                    if (TXT_RDT.Text == null || TXT_RDT.Text == "")
                    {
                        MessageBox.ShowMessage("반출일을 입력하세요.(0000년 00월 00일00:00 / 12자리) ", this.Page);
                        return;
                    }
                    else if (DDL_MAT.Text == null || DDL_MAT.Text == "")
                    {
                        MessageBox.ShowMessage("Scrap 종류를 선택하세요.", this.Page);
                        return;
                    }
                    else if (DDL_CUST.Text == null || DDL_CUST.Text == "")
                    {
                        MessageBox.ShowMessage("반출처를 선택하세요", this.Page);
                        return;
                    }
                    else if (TXT_SCRQTY.Text == null || TXT_SCRQTY.Text == "")
                    {
                        MessageBox.ShowMessage("Scrap 수량을 입력하세요.", this.Page);
                        return;
                    }

                    else if (TXT_DOCNO.Text == null || TXT_DOCNO.Text == "")
                    {
                        MessageBox.ShowMessage("반출증 No. 를 입력하세요", this.Page);
                        return;
                    }
                    else
                    {
                        string seq = dtInsert.Rows[r]["SEQ"].ToString();
                        string yn = dtInsert.Rows[r]["OUT_YN"].ToString() == "" ? "N" : dtInsert.Rows[r]["OUT_YN"].ToString();
                        string user = Session["User"].ToString();
                        /*이미 LOT_NO가 생성 안된경우*/
                        if (yn != "Y")
                        {
                            StringBuilder sbSQL = new StringBuilder();
                            sbSQL.Append("UPDATE OUT_MAT_DRAIN \n");
                            sbSQL.Append("set                      \n");
                            sbSQL.Append(" LOT_NO =   '" + lot_no + "' \n");
                            sbSQL.Append(",OUT_YN = 'Y'                \n");
                            sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
                            sbSQL.Append(",UPDT_DT  = GETDATE()             \n");
                            sbSQL.Append("where seq = '" + seq + "'             \n");

                            if (insertFlag == false && (r == dtInsert.Rows.Count-1))
                            {
                                sbSQL.Append("insert into OUT_MAT_HIS \n");
                                sbSQL.Append("(                      \n");
                                sbSQL.Append(" LOT_NO                \n");
                                sbSQL.Append(",STATE_FLAG            \n");
                                sbSQL.Append(",R_DT                  \n");
                                sbSQL.Append(",R_INPUT_DT            \n");
                                sbSQL.Append(",R_MAT                 \n");
                                sbSQL.Append(",R_QTY                 \n");
                                sbSQL.Append(",R_QTY_UNIT            \n");
                                sbSQL.Append(",R_CUST                \n");
                                sbSQL.Append(",R_DOC_NO              \n");
                                sbSQL.Append(",R_RMK                 \n");
                                sbSQL.Append(",IN_AU_QTY             \n");
                                sbSQL.Append(",IN_AU_UNIT            \n");
                                sbSQL.Append(",INSRT_USER_ID         \n");
                                sbSQL.Append(",INSRT_DT              \n");
                                sbSQL.Append(",UPDT_USER_ID          \n");
                                sbSQL.Append(",UPDT_DT               \n");
                                sbSQL.Append(")                      \n");
                                sbSQL.Append("VALUES(                \n");
                                sbSQL.Append("'" + lot_no + "'      \n");
                                sbSQL.Append(",'I'                    \n");
                                sbSQL.Append(",'" + r_dt + "'        \n");
                                sbSQL.Append(",'" + r_inputDt + "'   \n");
                                sbSQL.Append(",'" + r_mat + "'       \n");
                                sbSQL.Append(",'" + r_qty + "'       \n");
                                sbSQL.Append(",'" + r_unit + "'      \n");
                                sbSQL.Append(",'" + r_cust + "'      \n");
                                sbSQL.Append(",'" + r_docNo + "'     \n");
                                sbSQL.Append(",'" + r_rmk + "'       \n");
                                sbSQL.Append(",'" + AUQTY + "'       \n");
                                sbSQL.Append(",'" + AUUNIT + "'       \n");
                                sbSQL.Append(",'" + user + "'        \n");
                                sbSQL.Append(",GETDATE()             \n");
                                sbSQL.Append(",'" + user + "'        \n");
                                sbSQL.Append(",GETDATE()             \n");
                                sbSQL.Append(" )                     \n");
                                /*한번 입력 lot이냐 */
                                insertFlag = true;
                            }

                            string setSQL = sbSQL.ToString();
                            if (QueryExecute(setSQL) < 0)
                            {
                                MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                                return;
                            }
                            MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);
                        }
                    }
                }
            }


            for (int r = 0; r < dtDelete.Rows.Count; r++)
            {
            //취소 체크된 로직인지 확인하기
                if (dtDelete.Rows[r]["C_CHK"].ToString() == "1") //체크된 row이면
                {
                    string lotNo = dtDelete.Rows[r]["LOT_NO"].ToString();
                    string seq = dtDelete.Rows[r]["SEQ"].ToString();
                    string yn = dtDelete.Rows[r]["OUT_YN"].ToString() == "" ? "N" : dtDelete.Rows[r]["OUT_YN"].ToString();
                    string user = Session["User"].ToString();
                    /*이미 LOT_NO가 생성 안된경우*/
                    if (yn == "Y")
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        sbSQL.Append("UPDATE OUT_MAT_DRAIN \n");
                        sbSQL.Append("set                      \n");
                        sbSQL.Append(" LOT_NO =   NULL         \n");
                        sbSQL.Append(",OUT_YN = 'N'             \n");
                        sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
                        sbSQL.Append(",UPDT_DT  = GETDATE()             \n");
                        sbSQL.Append("where LOT_NO = '" + lotNo + "' \n");

                        sbSQL.Append("UPDATE OUT_MAT_HIS                \n");
                        sbSQL.Append("set                               \n");
                        sbSQL.Append(" STATE_FLAG =  'D'                \n");
                        sbSQL.Append(",UPDT_USER_ID  = '" + user + "'   \n");
                        sbSQL.Append(",UPDT_DT  = GETDATE()             \n");
                        sbSQL.Append("where LOT_NO = '" + lotNo + "' \n");

                        sbSQL.Append("DELETE OUT_MAT_HIS             \n");
                        sbSQL.Append("where LOT_NO = '" + lotNo + "' \n");

                        string setSQL = sbSQL.ToString();

                        if (QueryExecute(setSQL) < 0)
                        {
                            MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                            return;
                        }
                        MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);
                    }
                }
            }

        }

        public int QueryExecute(string sql)
        {
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = sql;

            try
            {
                value = sql_cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
                value = -1;
            }
            sql_conn.Close();
            return value;
        }


        protected void btn_mighty_save_Click(object sender, EventArgs e)
        {
            DataSet ds = FpSpread_new_data.DataSource as DataSet;

            FpSpread_new_data_UpdateCommand(ds.Tables[0]); ;
            
            search();
            //initTextBox();
            
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(CheckBox1.Checked == true)
            {
                SDDL_FAC.Enabled = false;
                DDL_MAT.Enabled = false;
                DDL_MACH.Enabled = false;
            }
            else
            {
                SDDL_FAC.Enabled = true;
                DDL_MAT.Enabled = true;
                DDL_MACH.Enabled = true;
            }
        }
        private void seqComBoSet(string FLAG )
        {
            /*공장 선택에 따른 콤보SET*/
            string plant = SDDL_FAC.Text;
            string mat = DDL_MAT.Text;

            /*SCRAP종류*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select DISTINCT GROUP1_CODE, GROUP1_DESC from dbo.SA_SYS_CODE \n");
            sbSQL.Append("where PLANT_CD = '" + plant + "' \n");

            DataTable MAT = fun.getData(sbSQL.ToString());
            if (MAT.Rows.Count > 0)
            {
                DDL_MAT.DataTextField = "GROUP1_DESC";
                DDL_MAT.DataValueField = "GROUP1_CODE";
                DDL_MAT.DataSource = MAT;
                DDL_MAT.DataBind();                
            }

            /*발생장비set*/
            StringBuilder sbSQL3 = new StringBuilder();
            sbSQL3.Append("select DISTINCT GROUP3_CODE, GROUP3_DESC from dbo.SA_SYS_CODE \n");
            sbSQL3.Append("where PLANT_CD = '" + plant + "' \n");
            /*scrap종류가 선택된게 있다면?*/
            if (FLAG == "FAC")
            {
                DDL_MAT.SelectedIndex = 0;
                mat = DDL_MAT.Text;
            }
            
            sbSQL3.Append("AND GROUP1_CODE = '" + mat + "' \n");


            DataTable MAC = fun.getData(sbSQL3.ToString());
            if (MAC.Rows.Count > 0)
            {
                DDL_MACH.DataTextField = "GROUP3_DESC";
                DDL_MACH.DataValueField = "GROUP3_CODE";
                DDL_MACH.DataSource = MAC;
                DDL_MACH.DataBind();
            }
        }

        protected void SDDL_FAC_SelectedIndexChanged(object sender, EventArgs e)
        {            
            seqComBoSet("FAC");            
        }

        protected void DDL_MAT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string mat = DDL_MAT.Text;            
            seqComBoSet("MAT");
            DDL_MACH.SelectedIndex = 0;
            DDL_MAT.Text = mat;
        }


        /*전체 취소부분*/
        protected void CHK_ALL_CheckedChanged(object sender, EventArgs e)
        {
            /*전체 조회 체크박스 선택시 chk못하게 lock*/
            if (CheckBox1.Checked == false)
            {
                DataSet ds = FpSpread_new_data.DataSource as DataSet;
                /*1차 검증 dataset 에 조회한 결과가 없을때*/
                if (ds == null || ds.Tables[0].Rows.Count <= 0)
                {
                    MessageBox.ShowMessage("전체 선택 대상이 없습니다. 조회를 먼저 하세요", this.Page);
                    return;
                }
                else
                {
                    /*2차 검증 LOT_NO 가 NULL인 데이터가 존재함 초기화를 위해 이리 처리함*/
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "DRAIN_IN_DT IS NOT NULL OR DRAIN_IN_DT <> ''";
                    DataTable dt = dv.Table;
                    if (dv.Count > 0)
                    {
                        for (int j = 0; j < FpSpread_new_data.Rows.Count; j++)
                        {
                            if (ds.Tables[0].Rows[j]["LOT_NO"] == null || ds.Tables[0].Rows[j]["LOT_NO"].ToString() == "")
                            {
                                if (CHK_ALL.Checked == true)
                                {
                                    ds.Tables[0].Rows[j][0] = 1;
                                }else
                                {
                                    ds.Tables[0].Rows[j][0] = 0;
                                }                                
                            }
                        }
                    }
                    else
                    {
                        MessageBox.ShowMessage("전체 선택 대상이 없습니다. 조회를 먼저 하세요", this.Page);
                        return;
                    }
                }
            }
        }

        protected void CHK_CNALL_CheckedChanged(object sender, EventArgs e)
        {
            /*전체 조회 체크박스 선택시 chk못하게 lock*/
            if (CheckBox1.Checked == false)
            {
                DataSet ds = FpSpread_new_data.DataSource as DataSet;
                /*1차 검증 dataset 에 조회한 결과가 없을때*/
                if (ds == null || ds.Tables[0].Rows.Count <= 0)
                {
                    MessageBox.ShowMessage("전체 선택 대상이 없습니다. 조회를 먼저 하세요", this.Page);
                    return;
                }
                else
                {
                    /*2차 검증 LOT_NO 가 NULL인 데이터가 존재함 초기화를 위해 이리 처리함*/
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "DRAIN_IN_DT IS NOT NULL OR DRAIN_IN_DT <> ''";
                    DataTable dt = dv.Table;
                    if (dv.Count > 0)
                    {
                        for (int j = 0; j < FpSpread_new_data.Rows.Count; j++)
                        {
                            if (ds.Tables[0].Rows[j]["LOT_NO"] != null && ds.Tables[0].Rows[j]["LOT_NO"].ToString() != "")
                            {
                                if (CHK_CNALL.Checked == true)
                                {
                                    ds.Tables[0].Rows[j][1] = 1;
                                }
                                else
                                {
                                    ds.Tables[0].Rows[j][1] = 0;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.ShowMessage("전체 선택 대상이 없습니다. 조회를 먼저 하세요", this.Page);
                        return;
                    }
                }
            }
        }
        /*선택 중량 합계*/
        protected void Button1_Click(object sender, EventArgs e)
        {
            DataSet ds = FpSpread_new_data.DataSource as DataSet;
            double sum = 0;
            double c_sum = 0;
            if (ds != null && ds.Tables[0].Rows.Count > 0 )
            {
                for (int i = 0; ds.Tables[0].Rows.Count > i; i++)
                {
                    if (ds.Tables[0].Rows[i]["CHK"].ToString() == "1")
                    {
                        double val = Convert.ToDouble(ds.Tables[0].Rows[i]["DRAIN_SCRAP_QTY"].ToString());
                        sum = sum + val;
                    }
                    if (ds.Tables[0].Rows[i]["C_CHK"].ToString() == "1")
                    {
                        double c_val = Convert.ToDouble(ds.Tables[0].Rows[i]["DRAIN_SCRAP_QTY"].ToString());
                        c_sum = c_sum + c_val;
                    }
                }
            }
            txt_sum.Text = sum.ToString("0.00");
            txt_c_sum.Text = c_sum.ToString("0.00");
        }

        protected void BUT_COMF3_Click(object sender, EventArgs e)
        {
            /*EXCEL저장*/
            string dt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");
            //FpSpread_new_data.SaveExcel("C:\\Scrap반출정보등록_" + dt + ".xls", FarPoint.Web.Spread.Model.IncludeHeaders.BothCustomOnly);
            //MessageBox.ShowMessage("Scrap반출정보등록[내컴퓨터 C:\\ 저장되었습니다.].", this.Page);
            System.IO.MemoryStream m_stream = new System.IO.MemoryStream();
            FpSpread_new_data.SaveExcel(m_stream, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders);
            m_stream.Position = 0;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "inline; filename=" + dt + ".xls");
            Response.BinaryWrite(m_stream.ToArray());
            Response.End();
        }       
        
    }
}
