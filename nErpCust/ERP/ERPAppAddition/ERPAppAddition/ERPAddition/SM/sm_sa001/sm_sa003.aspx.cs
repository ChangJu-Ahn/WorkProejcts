using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.SM.sm_sa001
{
    public partial class sm_sa003 : System.Web.UI.Page
    {
        sa_fun fun = new sa_fun();

        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();        
        DataSet ds = new DataSet();
        int value;
        string setSQL = "";


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

                DateTime setDate = DateTime.Today.AddDays(-7);
                tb_fr_yyyymmdd.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00") + setDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");

                FpSpread1.ClientAutoSize = false;
                FpSpread1.VerticalScrollBarPolicy = FarPoint.Web.Spread.ScrollBarPolicy.Always;                
            }
            grid();
            WebSiteCount();
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
            FpSpread1.Columns[0].DataField = "LOT_NO";
            FpSpread1.Columns[1].DataField = "R_DT";
            FpSpread1.Columns[2].DataField = "R_INPUT_DT";
            FpSpread1.Columns[3].DataField = "R_MAT";
            FpSpread1.Columns[4].DataField = "R_MAT_NM";
            FpSpread1.Columns[5].DataField = "R_QTY";
            FpSpread1.Columns[6].DataField = "R_QTY_UNIT";
            FpSpread1.Columns[7].DataField = "R_CUST";
            FpSpread1.Columns[8].DataField = "R_CUST_NM";
            FpSpread1.Columns[9].DataField = "R_DOC_NO";
            FpSpread1.Columns[10].DataField = "IN_AU_QTY";
            FpSpread1.Columns[11].DataField = "IN_AU_UNIT";
            FpSpread1.Columns[12].DataField = "R_RMK";            
            FpSpread1.Columns[13].DataField = "IN_DT";
            FpSpread1.Columns[14].DataField = "IN_QTY";
            FpSpread1.Columns[15].DataField = "IN_QTY_UNIT";
            FpSpread1.Columns[16].DataField = "IN_CUST";
            FpSpread1.Columns[17].DataField = "IN_CUST_NM";
            FpSpread1.Columns[18].DataField = "IN_RMK";

            FpSpread1.Columns[19].DataField = "SEND_DT";
            FpSpread1.Columns[20].DataField = "SEND_CUST";
            FpSpread1.Columns[21].DataField = "SEND_CUST_NM";
            FpSpread1.Columns[22].DataField = "SEND_QTY";            
            FpSpread1.Columns[23].DataField = "SEND_MAT";            
            FpSpread1.Columns[24].DataField = "SNED_YIELD";
            FpSpread1.Columns[25].DataField = "SNED_OUTQTY";
            FpSpread1.Columns[26].DataField = "R_MATE_QTY";
            FpSpread1.Columns[27].DataField = "SEND_RMK";
            
            FpSpread1.Columns[28].DataField = "ERP_PO_NO";
            FpSpread1.Columns[29].DataField = "ERP_PO_INDATE";
            FpSpread1.Columns[30].DataField = "ERP_PO_QTY";
            FpSpread1.Columns[31].DataField = "ERP_BAL";

            FpSpread1.Columns[32].DataField = "INSRT_DT";
            FpSpread1.Columns[33].DataField = "UPDT_DT";

            FpSpread1.Columns[0].Label = "LOT_NO";
            FpSpread1.Columns[1].Label = "반출일";
            FpSpread1.Columns[2].Label = "반출입력일";
            FpSpread1.Columns[3].Label = "반출품목CD";
            FpSpread1.Columns[3].Visible = false;
            FpSpread1.Columns[4].Label = "반출품목";
            FpSpread1.Columns[5].Label = "반출수량";
            FpSpread1.Columns[6].Label = "단위";
            FpSpread1.Columns[7].Label = "반출처CD";
            FpSpread1.Columns[7].Visible = false;
            FpSpread1.Columns[8].Label = "반출처";
            FpSpread1.Columns[9].Label = "반출증 번호";
            FpSpread1.Columns[10].Label = "Au농도";
            FpSpread1.Columns[11].Label = "단위";
            FpSpread1.Columns[12].Label = "비    고";

            FpSpread1.Columns[13].Label = "사급자재\n입고일";
            FpSpread1.Columns[14].Label = "수량";
            FpSpread1.Columns[15].Label = "단위";
            FpSpread1.Columns[16].Label = "입고자CD";
            FpSpread1.Columns[16].Visible = false;
            FpSpread1.Columns[17].Label = "입고자";
            FpSpread1.Columns[18].Label = "비    고";
            FpSpread1.Columns[18].Visible = false;
                        
            FpSpread1.Columns[19].Label = "사급자재 지급일";
            FpSpread1.Columns[20].Label = "지급처CD";
            FpSpread1.Columns[20].Visible = false;
            FpSpread1.Columns[21].Label = "지급처";
            FpSpread1.Columns[22].Label = "지급량";
            FpSpread1.Columns[23].Label = "원자재 종류";
            FpSpread1.Columns[24].Label = "수율";
            FpSpread1.Columns[25].Label = "임가공량";
            FpSpread1.Columns[26].Label = "원자재량";
            FpSpread1.Columns[27].Label = "사급자재 반출번호";            
       
            FpSpread1.Columns[28].Label = "사급자재 PO번호";
            FpSpread1.Columns[29].Label = "원자재 입고일";
            FpSpread1.Columns[30].Label = "원자재 입고량";
            FpSpread1.Columns[31].Label = "원자재 잔량";

            FpSpread1.Columns[32].Label = "입력일";
            FpSpread1.Columns[33].Label = "수정일";
            FpSpread1.Columns[32].Visible = false;
            FpSpread1.Columns[33].Visible = false;

            for (int c = 0; c < FpSpread1.Columns.Count; c++)
            {
                SetSpreadColumnLock(c);
            }            

            FpSpread1.ActiveSheetView.Columns[7].HorizontalAlign = HorizontalAlign.Left;
        }

        private void SetSpreadColumnLock(int column)
        {
            FpSpread1.ActiveSheetView.Protect = true;            
            FpSpread1.ActiveSheetView.LockForeColor = Color.Black;
            FpSpread1.ActiveSheetView.Columns[column].Font.Name = "돋움체";
            if (column > 1)
            {
                FpSpread1.ActiveSheetView.Columns[column].HorizontalAlign = HorizontalAlign.Center;
            }            
            FpSpread1.ActiveSheetView.Columns[column].VerticalAlign = VerticalAlign.Middle;
            FarPoint.Web.Spread.Column columnobj;

            int columncnt = FpSpread1.Columns.Count;

            columnobj = FpSpread1.ActiveSheetView.Columns[0, column]; // 입력된칼럼 lock
            columnobj.Locked = true;
            if (column == 10 || column ==11)
            {
                FpSpread1.ActiveSheetView.Columns[column].BackColor = Color.FromArgb(255, 204, 204);
            }
            else if (column < 13)
            {
                FpSpread1.ActiveSheetView.Columns[column].BackColor = Color.LightCyan;
            }

            else if (column < 18)
            {
                FpSpread1.ActiveSheetView.Columns[column].BackColor = Color.LightYellow;
            }
            else
            {
                FpSpread1.ActiveSheetView.Columns[column].BackColor = Color.FromArgb(204, 255, 204);
            }
        }



        private void setCombo()
        {
            /*수량단위 */
            DataTable UNIT = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8005'");
            if (UNIT.Rows.Count > 0)
            {
                DataRow dr = UNIT.NewRow();
                UNIT.Rows.InsertAt(dr, 0);

                DDL_INUNIT.DataTextField = "MINOR_NM";
                DDL_INUNIT.DataValueField = "MINOR_CD";
                DDL_INUNIT.DataSource = UNIT;
                DDL_INUNIT.DataBind();
            }

            /*반출처 */
            DataTable CUST = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004'");
            if (CUST.Rows.Count > 0)
            {
                DataTable ddt = CUST.Copy();
                DataRow dr = CUST.NewRow();
                CUST.Rows.InsertAt(dr, 0);

                DDL_INCUST.DataTextField = "MINOR_NM";
                DDL_INCUST.DataValueField = "MINOR_CD";
                DDL_INCUST.DataSource = CUST;
                DDL_INCUST.DataBind();

                /*사급자재 지급처*/
                DDL_SENDCUST.DataTextField = "MINOR_NM";
                DDL_SENDCUST.DataValueField = "MINOR_CD";
                DDL_SENDCUST.DataSource = CUST;
                DDL_SENDCUST.DataBind();
            }
            /*사급자재 종류 단위*/
            string[] SENDMAT = { "", "PGC", "Au Target" };
            DDL_SENDMAT.DataSource = SENDMAT;
            DDL_SENDMAT.DataBind();

            /*au농도 단위*/
            string[] AUUNIT = { "", "g/KG", "g/LT" };
            DDL_AUUNIT.DataSource = AUUNIT;
            DDL_AUUNIT.DataBind();
        }


        protected void FpSpread1_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            string row = e.CommandArgument.ToString();
            //{X=2,Y=3}
            row = row.Replace("{X=", "");
            row = row.Replace("Y=", "");
            string[] arryRow = row.Split(',');
            int rr = Convert.ToInt32(arryRow[0]);

            LOTNOBOX.Text = FpSpread1.Sheets[0].Cells[rr, 0].Text;            
            TXT_INDT.Text = FpSpread1.Sheets[0].Cells[rr, 13].Text;
            TXT_INQTY.Text = FpSpread1.Sheets[0].Cells[rr, 14].Text;
            DDL_INUNIT.Text = FpSpread1.Sheets[0].Cells[rr, 15].Text;
            TXT_SENDOUTUNIT.Text = FpSpread1.Sheets[0].Cells[rr, 15].Text;
            TXT_SENDQTYUNIT.Text = FpSpread1.Sheets[0].Cells[rr, 15].Text;
            DDL_INCUST.Text = FpSpread1.Sheets[0].Cells[rr, 16].Text;            
            TXT_SENDDT.Text = FpSpread1.Sheets[0].Cells[rr, 19].Text;
            DDL_SENDCUST.Text = FpSpread1.Sheets[0].Cells[rr, 20].Text;
            TXT_SENDQTY.Text = FpSpread1.Sheets[0].Cells[rr, 22].Text;
            DDL_SENDMAT.Text = FpSpread1.Sheets[0].Cells[rr, 23].Text;            
            TXT_SENDYIELD.Text = FpSpread1.Sheets[0].Cells[rr, 24].Text;
            TXT_SENDOUTQTY.Text = FpSpread1.Sheets[0].Cells[rr, 25].Text;            
            TXT_MATE_QTY.Text = FpSpread1.Sheets[0].Cells[rr, 26].Text;
            TXT_SENDRMK.Text = FpSpread1.Sheets[0].Cells[rr, 27].Text;

            TXT_AUQTY.Text = FpSpread1.Sheets[0].Cells[rr, 10].Text;
            DDL_AUUNIT.Text = FpSpread1.Sheets[0].Cells[rr, 11].Text;

            //DDL_PROC_TextChanged(null, null);
            /*선택된 row 색 칠하기*/
            for (int c = 0; c < FpSpread1.Rows.Count; c++)
            {
                FpSpread1.Rows[c].BackColor = Color.Empty;
            }
            FpSpread1.Rows[rr].BackColor = Color.LightPink;
        }

        protected void btn_mighty_retrieve0_Click(object sender, EventArgs e)
        {
            search();
            initTextBox();
        }

        private string getSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select		      \n");
            sbSQL.Append(" R.LOT_NO             \n");
            sbSQL.Append(",R.R_DT               \n");
            sbSQL.Append(",R.R_INPUT_DT         \n");
            sbSQL.Append(",R.R_MAT              \n");
            sbSQL.Append(",(SELECT TOP 1 MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8003' and MINOR_CD = R.R_MAT) AS R_MAT_NM /* 20151202 코드명으로 출력 이근만 SCRAP종류 */ \n");
            sbSQL.Append(",R.R_QTY              \n");
            sbSQL.Append(",R.R_QTY_UNIT         \n");
            sbSQL.Append(",R.R_CUST             \n");
            sbSQL.Append(",(SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004' AND MINOR_CD = R.R_CUST) AS R_CUST_NM /*20151202 코드명으로 출력 이근만 업체명*/ \n");
            sbSQL.Append(",R.R_DOC_NO           \n");
            sbSQL.Append(",R.R_RMK              \n");
            sbSQL.Append(",R.R_MATE_QTY         \n");
            sbSQL.Append(",R.IN_AU_QTY          \n");
            sbSQL.Append(",R.IN_AU_UNIT         \n");
            sbSQL.Append(",R.IN_DT              \n");
            sbSQL.Append(",CASE WHEN AU.AU_QTY IS NULL THEN R.IN_QTY ELSE AU.AU_QTY END IN_QTY             \n");
            sbSQL.Append(",CASE WHEN AU.WEIGHT_UNIT IS NULL THEN R.IN_QTY_UNIT ELSE CASE WHEN AU.WEIGHT_UNIT = 'g' THEN 'GR' ELSE AU.WEIGHT_UNIT END END IN_QTY_UNIT       \n");
            sbSQL.Append(",R.IN_CUST            \n");
            sbSQL.Append(",(SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004' AND MINOR_CD = R.IN_CUST) AS IN_CUST_NM /*20151202 코드명으로 출력 이근만 업체명*/ \n");
            sbSQL.Append(",R.IN_RMK             \n");
            sbSQL.Append(",R.SEND_MAT           \n");
            sbSQL.Append(",R.SEND_DT            \n");
            sbSQL.Append(",R.SEND_QTY           \n");
            sbSQL.Append(",R.SEND_CUST          \n");
            sbSQL.Append(",(SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8004' AND MINOR_CD = R.SEND_CUST) AS SEND_CUST_NM /*20151202 코드명으로 출력 이근만 업체명*/ \n");
            sbSQL.Append(",R.SNED_YIELD         \n");
            sbSQL.Append(",R.SNED_OUTQTY        \n");
            sbSQL.Append(",R.SEND_RMK           \n");
            sbSQL.Append(",R.ERP_PO_NO          \n");
            sbSQL.Append(",'' AS ERP_PO_INDATE       \n");
            sbSQL.Append(",'' AS ERP_PO_QTY          \n");
            sbSQL.Append(",'' AS ERP_BAL             \n");
            sbSQL.Append(",R.INSRT_DT           \n");
            sbSQL.Append(",R.UPDT_DT            \n");
            sbSQL.Append("from OUT_MAT_HIS R   \n");
            sbSQL.AppendLine(" LEFT OUTER JOIN OUT_LOT_AU_WEIGHT AU");
            sbSQL.AppendLine("   ON R.LOT_NO = AU.LOT_NO");
            sbSQL.AppendLine(" WHERE 1=1");
            sbSQL.AppendLine(" AND ISNULL(R_DOC_NO, '') <> '' ");
            /*입고대상 입고일이 없는것*/
            if (rdo01.Checked)
            {
                sbSQL.Append("AND  (R.IN_DT is null  or R.IN_DT='')    \n");
            }else if(rdo02.Checked)
            {
                /*지급대상 지급일이 없는것*/
                sbSQL.Append("where  R.IN_DT is R.not null    \n");
                sbSQL.Append("  and  (R.SEND_DT is null or R.SEND_DT = '')    \n");
            }

            sbSQL.Append("ORDER BY R.R_DT \n"); //16.07.19, AHNCJ : 반출일을 기준으로 정렬

            return sbSQL.ToString();
        }
        protected void search()
        {
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
                    FpSpread1.DataSource = ds.Tables["DataSet1"];
                    FpSpread1.DataBind();

                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                FpSpread1.DataSource = ds.Tables["DataSet1"];
                FpSpread1.DataBind();

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
        }


        protected void btn_mighty_save_Click(object sender, EventArgs e)
        {
            string LOTNO = LOTNOBOX.Text;            
            string INDT = TXT_INDT.Text;
            string INQTY = TXT_INQTY.Text == "" ? "0" : TXT_INQTY.Text;
            string INUNIT = DDL_INUNIT.Text;
            string INCUST = DDL_INCUST.Text;
            string MATE_QTY = TXT_MATE_QTY.Text == "" ? "0" : TXT_MATE_QTY.Text;

            string SENDMAT = DDL_SENDMAT.Text;
            string SENDDT = TXT_SENDDT.Text;
            string SENDCUST = DDL_SENDCUST.Text;
            string SENDQTY = TXT_SENDQTY.Text == "" ? "0" : TXT_SENDQTY.Text;
            string SENDYIELD = TXT_SENDYIELD.Text == "" ? "0" : TXT_SENDYIELD.Text;
            string SENDOUTQTY = TXT_SENDOUTQTY.Text == "" ? "0" : TXT_SENDOUTQTY.Text;
            string SENDRMK = TXT_SENDRMK.Text;

            string SENDAUQTY = TXT_AUQTY.Text == "" ? "0" : TXT_AUQTY.Text;
            string SENDAUUNIT = DDL_AUUNIT.Text;

            string user = Session["User"].ToString();          

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("update OUT_MAT_HIS \n");
            sbSQL.Append("set                      \n");
            sbSQL.Append(" STATE_FLAG = 'U' \n"); //상태값
            sbSQL.Append(",R_MATE_QTY = '" + MATE_QTY  + "' \n"); //원자재량
            
            sbSQL.Append(",IN_DT  = '" + INDT + "' \n");
            sbSQL.Append(",IN_QTY  = '" + INQTY + "' \n");
            sbSQL.Append(",IN_QTY_UNIT  = '" + INUNIT + "' \n");
            sbSQL.Append(",IN_CUST  = '" + INCUST + "' \n");

            sbSQL.Append(",SEND_MAT  = '" + SENDMAT + "' \n");
            sbSQL.Append(",SEND_DT  = '" + SENDDT + "' \n");
            sbSQL.Append(",SEND_QTY  = '" + SENDQTY + "' \n");
            sbSQL.Append(",SEND_CUST  = '" + SENDCUST + "' \n");
            sbSQL.Append(",SNED_YIELD  = '" + SENDYIELD + "' \n");
            sbSQL.Append(",SNED_OUTQTY  = '" + SENDOUTQTY + "' \n");
            sbSQL.Append(",SEND_RMK  = '" + SENDRMK + "' \n");

            sbSQL.Append(",IN_AU_QTY  = '" + SENDAUQTY + "' \n");
            sbSQL.Append(",IN_AU_UNIT  = '" + SENDAUUNIT + "' \n");

            sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
            sbSQL.Append(",UPDT_DT  =   GETDATE()               \n");
            sbSQL.Append("where LOT_NO ='" + LOTNO + "'      \n");
            setSQL = sbSQL.ToString();            

            if (QueryExecute(setSQL) < 0)
            {
                MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                return;
            }
            //initTextBox();            
            MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);
            search();
            //}
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

        public void initTextBox()
        {
            LOTNOBOX.Text = "";
            TXT_INDT.Text = "";
            TXT_INQTY.Text = "";
            DDL_INUNIT.Text = "";
            DDL_INCUST.Text = "";
            TXT_SENDDT.Text = "";
            DDL_SENDCUST.Text = "";
            TXT_SENDQTY.Text = "";
            TXT_SENDQTYUNIT.Text = "";
            DDL_SENDMAT.Text = "";
            TXT_SENDYIELD.Text = "";
            TXT_SENDOUTQTY.Text = "";
            TXT_SENDOUTUNIT.Text = "";
            TXT_MATE_QTY.Text = "";
            TXT_SENDRMK.Text = "";

            TXT_AUQTY.Text = "";
            DDL_AUUNIT.Text = "";
        }

        /*값변경시 계산 로직 호출*/
        protected void TXT_SENDQTY_TextChanged(object sender, EventArgs e)
        {
            setOutQty();
        }
        protected void TXT_SENDYIELD_TextChanged(object sender, EventArgs e)
        {
            setOutQty();
        }
        private void setOutQty()
        {
            double sendQty = Convert.ToDouble(TXT_SENDQTY.Text);
            double sendYield = Convert.ToDouble(TXT_SENDYIELD.Text);

            if(sendQty > 0 && sendYield > 0 )
            {
                double outQty = sendQty * (sendYield * 0.01);
                TXT_SENDOUTQTY.Text = outQty + "" ;
            }
            else
            {
                TXT_SENDOUTQTY.Text = "0";
            }
        }

        protected void TXT_SENDDT_TextChanged(object sender, EventArgs e)
        {
            if(TXT_INQTY.Text != "")
            {
                TXT_SENDQTY.Text = TXT_INQTY.Text;
                TXT_SENDQTYUNIT.Text = DDL_INUNIT.Text;
            }
        }

        protected void btn_mighty_excel_Click(object sender, EventArgs e)
        {
            string dt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");           
            System.IO.MemoryStream m_stream = new System.IO.MemoryStream();
            FpSpread1.SaveExcel(m_stream, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders);
            m_stream.Position = 0;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "inline; filename=" + dt + ".xls");
            Response.BinaryWrite(m_stream.ToArray());
            Response.End();
        }
    }
}