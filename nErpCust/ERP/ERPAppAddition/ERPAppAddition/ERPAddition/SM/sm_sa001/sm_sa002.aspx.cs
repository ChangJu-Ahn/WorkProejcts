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
using System.Windows.Forms;

namespace ERPAppAddition.ERPAddition.SM.sm_sa001
{
    public partial class sm_sa002 : System.Web.UI.Page
    {    
        sa_fun fun = new sa_fun();

        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string userid;       

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();        
        DataSet ds = new DataSet();
        int value;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                FpSpread1.ActiveSheetView.AutoPostBack = true;
                
                setCombo();

                DateTime setDate = DateTime.Today.AddDays(-7);
                tb_fr_yyyymmdd.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00") + setDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");

                grid();
                WebSiteCount();
            }
            if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                userid = "dev"; //erp에서 실행하지 않았을시 대비용
            else
                userid = Request.QueryString["userid"];
            Session["User"] = userid;

            //grid();
            //WebSiteCount();
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
            FpSpread1.Columns.Count = 24;
            FpSpread1.Rows.Count = 0;

            FpSpread1.Columns[0].DataField = "SEQ";
            FpSpread1.Columns[1].DataField = "STATE_FLAG";
            FpSpread1.Columns[1].Visible = false;
            FpSpread1.Columns[2].DataField = "DRAIN_IN_DT";
            FpSpread1.Columns[3].DataField = "DRAIN_PLANT";
            FpSpread1.Columns[3].Visible = false;
            FpSpread1.Columns[4].DataField = "DRAIN_PLANT_NM";
            FpSpread1.Columns[4].Width = 120;
            FpSpread1.Columns[5].DataField = "DRAIN_PROCESS";
            FpSpread1.Columns[6].DataField = "DRAIN_MACHINE";
            FpSpread1.Columns[7].DataField = "DRAIN_MAT";
            FpSpread1.Columns[7].Visible = false;
            FpSpread1.Columns[8].DataField = "DRAIN_MAT_NM";
            FpSpread1.Columns[8].Width = 120;
            FpSpread1.Columns[9].DataField = "DRAIN_QTY";
            FpSpread1.Columns[10].DataField = "DRAIN_SCRAP_QTY";
            FpSpread1.Columns[11].DataField = "DRAIN_UINT";
            FpSpread1.Columns[12].DataField = "DRAIN_INTGELEC";
            FpSpread1.Columns[13].DataField = "DRAIN_RMK";            

            FpSpread1.Columns[14].DataField = "COMF1_USER_ID";
            FpSpread1.Columns[15].DataField = "COMF1_DT";
            FpSpread1.Columns[16].DataField = "COMF2_USER_ID";
            FpSpread1.Columns[17].DataField = "COMF2_DT";

            FpSpread1.Columns[18].DataField = "DRAIN_ID";
            FpSpread1.Columns[19].DataField = "AMAT_IN";

            FpSpread1.Columns[20].DataField = "INSRT_USER_ID";
            FpSpread1.Columns[21].DataField = "INSRT_DT";
            FpSpread1.Columns[20].Visible = false;
            FpSpread1.Columns[21].Visible = false;

            FpSpread1.Columns[22].DataField = "UPDT_USER_ID";
            FpSpread1.Columns[23].DataField = "UPDT_DT";            
            
            FpSpread1.Columns[0].Label = "SEQ";
            FpSpread1.Columns[1].Label = "Flag";
            FpSpread1.Columns[2].Label = "발생일";
            FpSpread1.Columns[3].Label = "발생공장CD";
            FpSpread1.Columns[4].Label = "발생공장";
            FpSpread1.Columns[5].Label = "발생공정";
            FpSpread1.Columns[6].Label = "발생장비";
            FpSpread1.Columns[7].Label = "Scrap종류CD";
            FpSpread1.Columns[8].Label = "Scrap종류";
            FpSpread1.Columns[9].Label = "누적장수";
            FpSpread1.Columns[10].Label = "Scrap수량";
            FpSpread1.Columns[11].Label = "단위";
            FpSpread1.Columns[12].Label = "적산전력";
            FpSpread1.Columns[13].Label = "비고";

            FpSpread1.Columns[14].Label = "자재확인";
            FpSpread1.Columns[15].Label = "확인일";
            FpSpread1.Columns[16].Label = "환경확인";
            FpSpread1.Columns[17].Label = "확인일";

            FpSpread1.Columns[18].Label = "MCS ID";
            FpSpread1.Columns[19].Label = "투입량";

            FpSpread1.Columns[20].Label = "입력 USER";
            FpSpread1.Columns[21].Label = "입력일";
            FpSpread1.Columns[22].Label = "수정 USER";
            FpSpread1.Columns[23].Label = "수정일";

            for (int c = 0; c < FpSpread1.Columns.Count; c++)
            {
                SetSpreadColumnLock(c);
            }            
        }

        private void SetSpreadColumnLock(int column)
        {
            FpSpread1.ActiveSheetView.Protect = true;
            FpSpread1.ActiveSheetView.LockForeColor = Color.Black;
            FpSpread1.ActiveSheetView.Columns[column].BackColor = Color.LightCyan;
            FpSpread1.ActiveSheetView.Columns[column].HorizontalAlign = HorizontalAlign.Center;
            FpSpread1.ActiveSheetView.Columns[column].VerticalAlign = VerticalAlign.Middle;
            FpSpread1.ActiveSheetView.Columns[column].Font.Name = "돋움체";
            FarPoint.Web.Spread.Column columnobj;

            int columncnt = FpSpread1.Columns.Count;

            columnobj = FpSpread1.ActiveSheetView.Columns[0, column]; // 입력된칼럼 lock
            columnobj.Locked = true;

        }

        

        private void setCombo()
        {
            /*공장*/
            DataTable plant = fun.getData("select DISTINCT PLANT_CD, PLANT_DESC from dbo.SA_SYS_CODE");
            if (plant.Rows.Count > 0)
            {
                DDL_PLANT.DataTextField = "PLANT_DESC";
                DDL_PLANT.DataValueField = "PLANT_CD";
                DDL_PLANT.DataSource = plant;
                DDL_PLANT.DataBind();
            }

            /*공정*/
            DataTable PROCESS = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8001'");
            if (PROCESS.Rows.Count > 0)
            {
                DataRow dr = PROCESS.NewRow();
                PROCESS.Rows.InsertAt(dr, 0);

                DDL_PROC.DataTextField = "MINOR_NM";
                DDL_PROC.DataValueField = "MINOR_CD";
                DDL_PROC.DataSource = PROCESS;
                DDL_PROC.DataBind();
            }

            /*장비*/
            DataTable MAC = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8002'");
            if (MAC.Rows.Count > 0)
            {
                DataRow dr = MAC.NewRow();
                MAC.Rows.InsertAt(dr, 0);

                DDL_MACH.DataTextField = "MINOR_NM";
                DDL_MACH.DataValueField = "MINOR_CD";
                DDL_MACH.DataSource = MAC;
                DDL_MACH.DataBind();
            }
            /*SCRAP종류*/
            DataTable MAT = fun.getData("SELECT MINOR_CD, MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8003'");
            if (MAT.Rows.Count > 0)
            {
                DataRow dr = MAT.NewRow();
                MAT.Rows.InsertAt(dr, 0);

                DDL_MAT.DataTextField = "MINOR_NM";
                DDL_MAT.DataValueField = "MINOR_CD";
                DDL_MAT.DataSource = MAT;
                DDL_MAT.DataBind();
            }

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

        }


        protected void FpSpread1_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            string row = e.CommandArgument.ToString();
            //{X=2,Y=3}
            row = row.Replace("{X=", "");
            row = row.Replace("Y=", "");
            string[] arryRow = row.Split(',');
            int rr = Convert.ToInt32(arryRow[0]);

            DDL_MAT.Text = FpSpread1.Sheets[0].Cells[rr, 7].Text;
            SEQBOX.Text = FpSpread1.Sheets[0].Cells[rr, 0].Text;
            
            if (FpSpread1.Sheets[0].Cells[rr, 3].Text == "")
            {
                return;
            }
            DDL_PLANT.Text = FpSpread1.Sheets[0].Cells[rr, 3].Text;
            seqComBoSet();
            DDL_PROC.Text = FpSpread1.Sheets[0].Cells[rr, 5].Text;
            DDL_MACH.Text = FpSpread1.Sheets[0].Cells[rr, 6].Text;            
            TXT_QTY.Text = FpSpread1.Sheets[0].Cells[rr, 9].Text;
            TXT_SCRQTY.Text = FpSpread1.Sheets[0].Cells[rr, 10].Text;
            DDL_UNIT.Text = FpSpread1.Sheets[0].Cells[rr, 11].Text;
            
            TXT_RMK.Text = FpSpread1.Sheets[0].Cells[rr, 13].Text;
            
            DDL_MAT_SelectedIndexChanged(null, null);

            TXT_DRAINDT.Text = FpSpread1.Sheets[0].Cells[rr, 2].Text;
            TXT_IN_AMAT.Text = FpSpread1.Sheets[0].Cells[rr, 19].Text;
            TXT_INTGELEC.Text = FpSpread1.Sheets[0].Cells[rr, 12].Text;

            TXT_IN_AMAT.Enabled = false;
            TXT_INTGELEC.Enabled = false;
            //TXT_DRAINDT.Enabled = false;

            /*선택된 row 색 칠하기*/
            for (int c = 0; c < FpSpread1.Rows.Count; c++)
            {
                FpSpread1.Sheets[0].Rows[c].BackColor = Color.Empty;
            }
            FpSpread1.Sheets[0].Rows[rr].BackColor = Color.LightPink;
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
            sbSQL.Append("select             \n");
            sbSQL.Append(" OT.seq			\n");
            sbSQL.Append(",OT.STATE_FLAG			\n");
            sbSQL.Append(",OT.DRAIN_IN_DT			\n");
            sbSQL.Append(",OT.DRAIN_PLANT        \n");
            sbSQL.Append(",(select TOP 1 PLANT_DESC from SA_SYS_CODE WHERE PLANT_CD = OT.DRAIN_PLANT) AS DRAIN_PLANT_NM  /*20151229 공장 명으로 변경  기존 : DRAIN_PLANT*/ \n");
            sbSQL.Append(",OT.DRAIN_PROCESS      \n");
            sbSQL.Append(",OT.DRAIN_MACHINE      \n");
            sbSQL.Append(",OT.DRAIN_MAT          \n");
            sbSQL.Append(",(SELECT TOP 1 MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'S8003' and MINOR_CD = OT.DRAIN_MAT ) AS DRAIN_MAT_NM /*코드 명으로 넣어달라 요청 20151229 이근만*/ \n");
            sbSQL.Append(",OT.DRAIN_QTY          \n");
            sbSQL.Append(",OT.DRAIN_SCRAP_QTY    \n");
            sbSQL.Append(",OT.DRAIN_UINT         \n");
            sbSQL.Append(",OT.DRAIN_INTGELEC     \n");
            sbSQL.Append(",OT.DRAIN_RMK          \n");
            sbSQL.Append(",OT.COMF1_USER_ID     \n");
            sbSQL.Append(",OT.COMF1_DT          \n");
            sbSQL.Append(",OT.COMF2_USER_ID     \n");
            sbSQL.Append(",OT.COMF2_DT          \n");
            sbSQL.Append(",OT.DRAIN_ID          \n");
            sbSQL.Append(",OT.AMAT_IN          \n");
            sbSQL.Append(",OT.INSRT_USER_ID      \n");
            sbSQL.Append(",OT.INSRT_DT           \n");
            sbSQL.Append(",OT.UPDT_USER_ID       \n");
            sbSQL.Append(",OT.UPDT_DT            \n");
            sbSQL.Append("from OUT_MAT_DRAIN OT \n");
            sbSQL.Append("where (OT.OUT_YN IS NULL OR OT.OUT_YN = 'N') AND OT.STATE_FLAG <> 'D' \n");
            sbSQL.Append("  AND SUBSTRING(OT.DRAIN_IN_DT, 1, 8) BETWEEN '" + tb_fr_yyyymmdd.Text + "' AND '" + tb_to_yyyymmdd.Text + "' \n");
            sbSQL.Append("  AND OT.DRAIN_PLANT = '" + DDL_PLANT.Text + "' \n");
            sbSQL.Append("  ORDER BY OT.SEQ DESC\n");
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

        protected void btn_mighty_insert_Click(object sender, EventArgs e)
        {
            initTextBox();
            SEQBOX.Text = "신규추가";
            TXT_DRAINDT.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00");
        }

        protected void btn_mighty_save_Click(object sender, EventArgs e)
        {
            string SEQ = SEQBOX.Text;
            string DRAINDT = TXT_DRAINDT.Text;
            string PLANT = DDL_PLANT.Text;
            string PROC = DDL_PROC.Text;
            string MACH = DDL_MACH.Text;
            string MAT = DDL_MAT.Text;
            string QTY = TXT_QTY.Text;
            string MCSLOT = DDL_MCSLOT_ID.Text;
            string INAMAT = TXT_IN_AMAT.Text;

            if(QTY =="" || QTY == null)
            {
                QTY = "0";
            }
            string SCRQTY = TXT_SCRQTY.Text;
            if (SCRQTY == "" || SCRQTY == null)
            {
                SCRQTY = "0";
            }

            if (INAMAT == "" || INAMAT == null)
            {
                INAMAT = "0";
            }

            string UNIT = DDL_UNIT.Text;
            string INTGELEC = TXT_INTGELEC.Text;
            string RMK = TXT_RMK.Text;

            string user = Session["User"].ToString();            

            if (DRAINDT == null || DRAINDT == "" || DRAINDT.Length  < 12)
                MessageBox.ShowMessage("발생일시를 입력하세요(12자)[yyyymmddhh24mi].", this.Page);
            else if (PLANT == null || PLANT == "")
                MessageBox.ShowMessage("공장 정보를 선택하세요.", this.Page);
            else if (PROC == null || PROC == "")
                MessageBox.ShowMessage("발생 공정을 선택하세요", this.Page);
            else if (MACH == null || MACH == "")
                MessageBox.ShowMessage("발생 장비를 선택하세요.", this.Page);
            else if (MAT == null || MAT == "")
                MessageBox.ShowMessage("Scrap 종류를 선택하세요.", this.Page);            
            else if (SCRQTY == null || SCRQTY == "")
                MessageBox.ShowMessage("수량을 입력하세요.", this.Page);
            else if (UNIT == null || UNIT == "")
                MessageBox.ShowMessage("수량의 단위를 선택하세요.", this.Page);
            else if (UNIT == null || UNIT == "")
                MessageBox.ShowMessage("수량의 단위를 선택하세요.", this.Page);
            else if (PROC == "PLAT" && MCSLOT == "")
                MessageBox.ShowMessage("Drain ID를 선택해주세요.", this.Page);
            else if (PROC == "SPUTTER" && MCSLOT == "")
                MessageBox.ShowMessage("Target ID를 선택해주세요.", this.Page); 

            else
            {
                string setSQL = "";
                if (SEQ == "신규추가" || SEQ == "")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("insert into OUT_MAT_DRAIN \n");
                    sbSQL.Append("(                      \n");
                    sbSQL.Append(" STATE_FLAG            \n");
                    sbSQL.Append(",DRAIN_IN_DT           \n");
                    sbSQL.Append(",DRAIN_PLANT           \n");
                    sbSQL.Append(",DRAIN_PROCESS         \n");
                    sbSQL.Append(",DRAIN_MACHINE         \n");
                    sbSQL.Append(",DRAIN_MAT             \n");
                    sbSQL.Append(",DRAIN_QTY             \n");
                    sbSQL.Append(",DRAIN_SCRAP_QTY       \n");
                    sbSQL.Append(",DRAIN_UINT            \n");
                    sbSQL.Append(",DRAIN_INTGELEC        \n");
                    sbSQL.Append(",DRAIN_RMK             \n");

                    sbSQL.Append(",COMF1_USER_ID         \n");
                    sbSQL.Append(",COMF1_DT              \n");
                    sbSQL.Append(",COMF2_USER_ID         \n");
                    sbSQL.Append(",COMF2_DT              \n");

                    sbSQL.Append(",DRAIN_ID              \n");
                    sbSQL.Append(",AMAT_IN              \n");

                    sbSQL.Append(",INSRT_USER_ID         \n");
                    sbSQL.Append(",INSRT_DT              \n");
                    sbSQL.Append(",UPDT_USER_ID          \n");
                    sbSQL.Append(",UPDT_DT               \n");
                    sbSQL.Append(")                      \n");
                    sbSQL.Append("VALUES(                \n");
                    sbSQL.Append("'I'                     \n");
                    sbSQL.Append(",'" + DRAINDT + "'      \n");
                    sbSQL.Append(",'" + PLANT + "'       \n");
                    sbSQL.Append(",'" + PROC + "'        \n");
                    sbSQL.Append(",'" + MACH + "'        \n");
                    sbSQL.Append(",'" + MAT + "'         \n");
                    sbSQL.Append(",'" + QTY + "'         \n");
                    sbSQL.Append(",'" + SCRQTY + "'      \n");
                    sbSQL.Append(",'" + UNIT + "'        \n");
                    sbSQL.Append(",'" + INTGELEC + "'    \n");
                    sbSQL.Append(",'" + RMK + "'         \n");

                    sbSQL.Append(",''                    \n");
                    sbSQL.Append(",''                    \n");
                    sbSQL.Append(",''                    \n");
                    sbSQL.Append(",''                    \n");
                    
                    sbSQL.Append(",'" + MCSLOT + "'      \n");
                    sbSQL.Append(",'" + INAMAT + "'      \n");

                    sbSQL.Append(",'" + user + "'        \n");
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(",'" + user + "'        \n");
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(" )                     \n");
                    setSQL = sbSQL.ToString();
                }
                else
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("update OUT_MAT_DRAIN \n");
                    sbSQL.Append("set                      \n");
                    sbSQL.Append(" STATE_FLAG = 'U' \n");
                    sbSQL.Append(",DRAIN_IN_DT = '" + DRAINDT + "' \n");
                    sbSQL.Append(",DRAIN_PLANT = '" + PLANT + "' \n");
                    sbSQL.Append(",DRAIN_PROCESS  = '" + PROC + "' \n");
                    sbSQL.Append(",DRAIN_MACHINE  = '" + MACH + "' \n");
                    sbSQL.Append(",DRAIN_MAT  = '" + MAT + "' \n");
                    sbSQL.Append(",DRAIN_QTY  = '" + QTY + "' \n");
                    sbSQL.Append(",DRAIN_SCRAP_QTY  = '" + SCRQTY + "' \n");
                    sbSQL.Append(",DRAIN_UINT  = '" + UNIT + "' \n");
                    sbSQL.Append(",DRAIN_INTGELEC  = '" + INTGELEC + "' \n");
                    sbSQL.Append(",DRAIN_RMK  = '" + RMK + "' \n");

                    sbSQL.Append(",DRAIN_ID  = '" + MCSLOT + "' \n");
                    sbSQL.Append(",AMAT_IN  = '" + INAMAT + "' \n");

                    sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
                    sbSQL.Append(",UPDT_DT  = GETDATE()               \n");
                    sbSQL.Append("where SEQ ='" + SEQ + "'      \n");
                    setSQL = sbSQL.ToString();
                }

                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }
                //initTextBox();
                search();
                MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);
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

        public void initTextBox()
        {
            SEQBOX.Text = "";
            TXT_DRAINDT.Text = "";
            DDL_PROC.Text = "";
            DDL_MACH.Text = "";
            DDL_MAT.Text = "";
            TXT_QTY.Text = "";
            TXT_SCRQTY.Text = "";
            DDL_UNIT.Text = "";
            TXT_INTGELEC.Text = "";
            TXT_RMK.Text = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("DRAIN_ID");

            DataRow dr = dt.NewRow();
            dt.Rows.InsertAt(dr, 0);

            DDL_MCSLOT_ID.DataTextField = "DRAIN_ID";
            DDL_MCSLOT_ID.DataValueField = "DRAIN_ID";


            DDL_MCSLOT_ID.Text = "";
            DDL_MCSLOT_ID.DataSource = dt;
            DDL_MCSLOT_ID.DataBind();

            TXT_IN_AMAT.Text = "";

            TXT_INTGELEC.Enabled = true;
            TXT_DRAINDT.Enabled = true;

            lblAmatIn.Visible = false;
            TXT_IN_AMAT.Visible = false;

            lblLotID.Visible = false;
            DDL_MCSLOT_ID.Visible = false;
        }

        protected void btn_mighty_delete_Click(object sender, EventArgs e)
        {
            string user = Session["User"].ToString();            
            string SEQ = SEQBOX.Text;
            StringBuilder sbSQL = new StringBuilder();                        
            sbSQL.Append("update OUT_MAT_DRAIN \n");
            sbSQL.Append("set                      \n");
            sbSQL.Append(" STATE_FLAG = 'D' \n");
            sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
            sbSQL.Append(",UPDT_DT  = GETDATE()               \n");
            sbSQL.Append("where SEQ ='" + SEQ + "' \n");
            QueryExecute(sbSQL.ToString());
            initTextBox();
            MessageBox.ShowMessage("정보가 삭제되었습니다.", this.Page);
            search();            
        }


        protected void BUT_COMF1_Click(object sender, EventArgs e)
        {
            if (SEQBOX.Text != "")
            {
                string user = Session["User"].ToString();
                string SEQ = SEQBOX.Text;
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("update OUT_MAT_DRAIN \n");
                sbSQL.Append("set                      \n");
                sbSQL.Append(" STATE_FLAG = 'U' \n");
                sbSQL.Append(",COMF1_USER_ID  = '" + user + "' \n");
                sbSQL.Append(",COMF1_DT  = CONVERT(VARCHAR, GETDATE(), 112)               \n");
                sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
                sbSQL.Append(",UPDT_DT  = GETDATE()               \n");
                sbSQL.Append("where SEQ ='" + SEQ + "' \n");
                QueryExecute(sbSQL.ToString());
                search();
                MessageBox.ShowMessage("확인정보가 등록 되었습니다.", this.Page);
            }
            else
            {
                MessageBox.ShowMessage("확인 대상을 먼저 선택하세요", this.Page);
            }
            
        }

        protected void BUT_COMF2_Click(object sender, EventArgs e)
        {
            if (SEQBOX.Text != "")
            {
                string user = Session["User"].ToString();
                string SEQ = SEQBOX.Text;
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("update OUT_MAT_DRAIN \n");
                sbSQL.Append("set                      \n");
                sbSQL.Append(" STATE_FLAG = 'U' \n");
                sbSQL.Append(",COMF2_USER_ID  = '" + user + "' \n");
                sbSQL.Append(",COMF2_DT  = CONVERT(VARCHAR, GETDATE(), 112)               \n");
                sbSQL.Append(",UPDT_USER_ID  = '" + user + "' \n");
                sbSQL.Append(",UPDT_DT  = GETDATE()               \n");
                sbSQL.Append("where SEQ ='" + SEQ + "' \n");
                QueryExecute(sbSQL.ToString());
                search();
                MessageBox.ShowMessage("확인정보가 등록 되었습니다.", this.Page);
            }
            else
            {
                MessageBox.ShowMessage("확인 대상을 먼저 선택하세요", this.Page);
            }
        }

        protected void DDL_MAT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string buMat = DDL_MAT.Text;
            if(DDL_MAT.Text == "S05")
            {
                DDL_UNIT.Text = "EA";
            }
            else
            {
                DDL_UNIT.Text = "KG";
            }

            if (DDL_MAT.Text == "S01" || DDL_MAT.Text == "S02" || DDL_MAT.Text == "S03" || DDL_MAT.Text == "S04")
            {
                TXT_QTY.Enabled = true;
                TXT_QTY.BackColor = Color.FromName("#FFFFCC");                
            }
            else
            {
                TXT_QTY.Enabled = false;
                TXT_QTY.BackColor = Color.Empty;
                TXT_QTY.Text = "";
            }

            if (DDL_MAT.Text == "S01")
            {
                lblLotID.Visible = true;
                DDL_MCSLOT_ID.Visible = true;

                lblAmatIn.Visible = true;
                TXT_IN_AMAT.Visible = true;

                TXT_DRAINDT.Enabled = false;

                lblLotID.Text = "Drain ID";
                lblLFT_Unit.Text = "mA";

                DataTable dt = new DataTable();
                dt.Columns.Add("DRAIN_ID");

                DataRow dr = dt.NewRow();
                dt.Rows.InsertAt(dr, 0);

                DDL_MCSLOT_ID.DataTextField = "DRAIN_ID";
                DDL_MCSLOT_ID.DataValueField = "DRAIN_ID";

                DDL_MCSLOT_ID.DataSource = dt;
                DDL_MCSLOT_ID.DataBind();

                TXT_IN_AMAT.Text = "";
                TXT_INTGELEC.Text = "";
            }
            else if (DDL_MAT.Text == "S03" || DDL_MAT.Text == "S04")
            {
                lblLotID.Visible = true;
                DDL_MCSLOT_ID.Visible = true;

                lblAmatIn.Visible = true;
                TXT_IN_AMAT.Visible = true;

                lblLotID.Text = "Target ID";
                lblLFT_Unit.Text = "Kwh";

                if (!TXT_DRAINDT.Enabled)
                {
                    TXT_DRAINDT.Text = "";
                }

                TXT_INTGELEC.Enabled = true;
                TXT_DRAINDT.Enabled = true;


                DataTable dt = new DataTable();
                dt.Columns.Add("MAT_LOT_NO");

                DataRow dr = dt.NewRow();
                dt.Rows.InsertAt(dr, 0);

                DDL_MCSLOT_ID.DataTextField = "MAT_LOT_NO";
                DDL_MCSLOT_ID.DataValueField = "MAT_LOT_NO";

                DDL_MCSLOT_ID.DataSource = dt;
                DDL_MCSLOT_ID.DataBind();

                TXT_IN_AMAT.Text = "";
                TXT_INTGELEC.Text = "";
            }
            else
            {
                lblLotID.Visible = false;
                DDL_MCSLOT_ID.Visible = false;

                lblAmatIn.Visible = false;
                TXT_IN_AMAT.Visible = false;

                TXT_INTGELEC.Enabled = true;
                if (!TXT_DRAINDT.Enabled)
                {
                    TXT_DRAINDT.Text = "";
                }
                TXT_DRAINDT.Enabled = true;

                TXT_IN_AMAT.Text = "";
                TXT_INTGELEC.Text = "";
            }

            seqComBoSet();
            DDL_MAT.Text = buMat;
        }

        protected void DDL_PLANT_SelectedIndexChanged(object sender, EventArgs e)
        {
            DDL_MAT.Text = "";
            seqComBoSet();
        }

        private void seqComBoSet()
        {
            /*공장 선택에 따른 콤보SET*/
            string plant = DDL_PLANT.Text;
            string mat = DDL_MAT.Text;

            /*SCRAP종류*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select DISTINCT GROUP1_CODE, GROUP1_DESC from dbo.SA_SYS_CODE \n");
            sbSQL.Append("where PLANT_CD = '" + plant + "' \n");

            DataTable MAT = fun.getData(sbSQL.ToString());
            if (MAT.Rows.Count > 0)
            {
                DataRow dr = MAT.NewRow();
                MAT.Rows.InsertAt(dr, 0);

                DDL_MAT.DataTextField = "GROUP1_DESC";
                DDL_MAT.DataValueField = "GROUP1_CODE";
                DDL_MAT.DataSource = MAT;
                DDL_MAT.DataBind();
            }

            /*발생공정종류*/
            StringBuilder sbSQL2 = new StringBuilder();
            sbSQL2.Append("select DISTINCT GROUP2_CODE, GROUP2_DESC from dbo.SA_SYS_CODE \n");
            sbSQL2.Append("where PLANT_CD = '" + plant + "' \n");
            /*scrap종류가 선택된게 있다면?*/
            if (mat != "" || mat == null)
            {
                sbSQL2.Append("AND GROUP1_CODE = '" + mat + "' \n");
            }

            DataTable PROCESS = fun.getData(sbSQL2.ToString());
            if (PROCESS.Rows.Count > 0)
            {
                DataRow dr = PROCESS.NewRow();
                PROCESS.Rows.InsertAt(dr, 0);

                DDL_PROC.DataTextField = "GROUP2_DESC";
                DDL_PROC.DataValueField = "GROUP2_CODE";
                DDL_PROC.DataSource = PROCESS;
                DDL_PROC.DataBind();
            }

            /*발생장비set*/
            StringBuilder sbSQL3 = new StringBuilder();
            sbSQL3.Append("select DISTINCT GROUP3_CODE, GROUP3_DESC from dbo.SA_SYS_CODE \n");
            sbSQL3.Append("where PLANT_CD = '" + plant + "' \n");
            /*scrap종류가 선택된게 있다면?*/
            if (mat != "" || mat == null)
            {
                sbSQL3.Append("AND GROUP1_CODE = '" + mat + "' \n");
            }

            DataTable MAC = fun.getData(sbSQL3.ToString());
            if (MAC.Rows.Count > 0)
            {
                DataRow dr = MAC.NewRow();
                MAC.Rows.InsertAt(dr, 0);

                DDL_MACH.DataTextField = "GROUP3_DESC";
                DDL_MACH.DataValueField = "GROUP3_CODE";
                DDL_MACH.DataSource = MAC;
                DDL_MACH.DataBind();
            }
        }

        protected void FpSpread1_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {

        }

        protected void BUT_COMF3_Click(object sender, EventArgs e)
        {
            string dt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00")+"_" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");
            //FpSpread1.SaveExcel("C:\\Scrap_"+dt+".xlsx", FarPoint.Web.Spread.Model.IncludeHeaders.BothCustomOnly);            
            //MessageBox.ShowMessage("Scrap발생정보등록[내컴퓨터 C:\\ 저장되었습니다.].", this.Page);

            System.IO.MemoryStream m_stream = new System.IO.MemoryStream();
            FpSpread1.SaveExcel(m_stream, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders);
            m_stream.Position = 0;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "inline; filename=" + dt + ".xls");
            Response.BinaryWrite(m_stream.ToArray());
            Response.End();

        }

        protected void DDL_MACH_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDL_MAT.Text == "S01")
            {

                if (DDL_MACH.SelectedIndex > 0)
                {
                    /*Drain ID set*/
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.AppendLine(" SELECT ");
                    sbSQL.AppendLine(" B.DRAIN_ID ");
                    sbSQL.AppendLine(" , B.ITEM_QTY ");
                    sbSQL.AppendLine(" , B.LIFE_TIME ");
                    sbSQL.AppendLine(" FROM OPENQUERY(CCUBE, ' ");
                    sbSQL.AppendLine(" SELECT ");
                    sbSQL.AppendLine("     CASE WHEN EQUIPMENT LIKE ''%-%'' THEN EQUIPMENT||''-''|| TRIM(TO_CHAR(TO_NUMBER(SUBSTR(MODULE,5)), ''09''))");
                    sbSQL.AppendLine("           ELSE EQUIPMENT END EQP_ID");
                    sbSQL.AppendLine("      ,MAT_BATCH_ID DRAIN_ID");
                    sbSQL.AppendLine("      ,ITEM_QTY ");
                    sbSQL.AppendLine("      ,LIFE_TIME");
                    sbSQL.AppendLine(" FROM AMAT_EQPMAT_BATCH_STS A");
                    sbSQL.AppendLine(" WHERE PLANT = ''CCUBEDIGITAL''");
                    sbSQL.AppendLine(" AND STATUS = ''DRAIN'' ");
                    sbSQL.AppendLine(" AND DRAIN_DATE > ''2016'' ");
                    sbSQL.AppendLine(" ') B ");
                    sbSQL.AppendLine(" WHERE 1=1");
                    sbSQL.AppendLine(" AND NOT EXISTS (SELECT * FROM OUT_MAT_DRAIN A WHERE A.DRAIN_ID = B.DRAIN_ID  AND STATE_FLAG <> 'D' ) ");

                    sbSQL.AppendLine(" AND B.EQP_ID = '" + DDL_MACH.Text + "'");


                    sbSQL.AppendLine(" ORDER BY DRAIN_ID ");


                    DataTable dt = fun.getData(sbSQL.ToString());
                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.InsertAt(dr, 0);

                        DDL_MCSLOT_ID.DataTextField = "DRAIN_ID";
                        DDL_MCSLOT_ID.DataValueField = "DRAIN_ID";
                        
                    }
                    else
                    {
                        dt = new DataTable();
                        dt.Columns.Add("DRAIN_ID");

                        DataRow dr = dt.NewRow();
                        dt.Rows.InsertAt(dr, 0);

                    }
                    DDL_MCSLOT_ID.DataTextField = "DRAIN_ID";
                    DDL_MCSLOT_ID.DataValueField = "DRAIN_ID";

                    DDL_MCSLOT_ID.DataSource = dt;
                    DDL_MCSLOT_ID.DataBind();
                }
            }
            else if (DDL_MAT.Text == "S03")
            {
                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine(" 	* ");
                sbSQL.AppendLine("  FROM OPENQUERY(CCUBE, ' ");
                sbSQL.AppendLine("   SELECT" );
                sbSQL.AppendLine("   INPUT_TIME" );
                sbSQL.AppendLine("   , LEAD(INPUT_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS OUT_TIME ");
                sbSQL.AppendLine("   , MAT_LOT_NO ");
                sbSQL.AppendLine("   , ITEM_CODE ");
                sbSQL.AppendLine("   , ITEM_QTY ");
                sbSQL.AppendLine("   , LIFE_TIME "); 
                sbSQL.AppendLine("   , EQUIPMENT  ");
                sbSQL.AppendLine("   , LEAD(LIFE_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS LIFE_TIME1  ");
                sbSQL.AppendLine("   , NUM  ");
                sbSQL.AppendLine(" FROM  ");
                sbSQL.AppendLine(" (  ");
                sbSQL.AppendLine("   SELECT   ");
                sbSQL.AppendLine("     ROW_NUMBER () OVER (PARTITION BY EQUIPMENT ORDER BY TRANS_TIME) AS NUM  ");
                sbSQL.AppendLine("     , SUBSTR(A.TRANS_TIME, 0, 12) AS INPUT_TIME ");
                sbSQL.AppendLine("     , A.MAT_LOT_NO ");
                sbSQL.AppendLine("     , A.ITEM_CODE ");
                sbSQL.AppendLine("     , A.ITEM_QTY ");
                sbSQL.AppendLine("     , A.LIFE_TIME ");
                sbSQL.AppendLine("     , A.EQUIPMENT ");
                sbSQL.AppendLine("   FROM AMAT_EQPMAT_HISTORY A  ");
                sbSQL.AppendLine("   WHERE PLANT = ''CCUBEDIGITAL''  ");
                sbSQL.AppendLine("   AND TRANSACTION  = ''MEMC''  ");
                sbSQL.AppendLine("   AND EQUIPMENT LIKE ''SPU%'' ");
                sbSQL.AppendLine("   AND ITEM_CODE IN (''1B30-02008'', ''1130-00031'')  ");
                sbSQL.AppendLine("   AND INTERLOCK_ID IS NULL ");
                sbSQL.AppendLine("   AND TRANS_TIME > ''20161201'' ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" ')B ");
                sbSQL.AppendLine(" WHERE EQUIPMENT = '" + DDL_MACH.Text + "'" );
                sbSQL.AppendLine(" AND NOT EXISTS (SELECT * FROM OUT_MAT_DRAIN A WHERE A.DRAIN_ID = B.MAT_LOT_NO  AND STATE_FLAG <> 'D' AND DRAIN_MAT = 'S03' AND DRAIN_MACHINE = EQUIPMENT ) ");
                sbSQL.AppendLine(" ORDER BY INPUT_TIME DESC ");
                
                DataTable dt = fun.getData(sbSQL.ToString());
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.NewRow();
                    dt.Rows.InsertAt(dr, 0);

                    DDL_MCSLOT_ID.DataTextField = "MAT_LOT_NO";
                    DDL_MCSLOT_ID.DataValueField = "MAT_LOT_NO";
                        
                }
                else
                {
                    dt = new DataTable();
                    dt.Columns.Add("MAT_LOT_NO");

                    DataRow dr = dt.NewRow();
                    dt.Rows.InsertAt(dr, 0);

                }
                DDL_MCSLOT_ID.DataTextField = "MAT_LOT_NO";
                DDL_MCSLOT_ID.DataValueField = "MAT_LOT_NO";

                DDL_MCSLOT_ID.DataSource = dt;
                DDL_MCSLOT_ID.DataBind();
                
            }
            else if (DDL_MAT.Text == "S04")
            {
                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine(" 	* ");
                sbSQL.AppendLine("  FROM OPENQUERY(CCUBE, ' ");
                sbSQL.AppendLine("   SELECT");
                sbSQL.AppendLine("   INPUT_TIME");
                sbSQL.AppendLine("   , LEAD(INPUT_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS OUT_TIME ");
                sbSQL.AppendLine("   , MAT_LOT_NO ");
                sbSQL.AppendLine("   , ITEM_CODE ");
                sbSQL.AppendLine("   , ITEM_QTY ");
                sbSQL.AppendLine("   , LIFE_TIME ");
                sbSQL.AppendLine("   , EQUIPMENT  ");
                sbSQL.AppendLine("   , LEAD(LIFE_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS LIFE_TIME1  ");
                sbSQL.AppendLine("   , NUM  ");
                sbSQL.AppendLine(" FROM  ");
                sbSQL.AppendLine(" (  ");
                sbSQL.AppendLine("   SELECT   ");
                sbSQL.AppendLine("     ROW_NUMBER () OVER (PARTITION BY EQUIPMENT ORDER BY TRANS_TIME) AS NUM  ");
                sbSQL.AppendLine("     , SUBSTR(A.TRANS_TIME, 0, 12) AS INPUT_TIME ");
                sbSQL.AppendLine("     , A.MAT_LOT_NO ");
                sbSQL.AppendLine("     , A.ITEM_CODE ");
                sbSQL.AppendLine("     , A.ITEM_QTY ");
                sbSQL.AppendLine("     , A.LIFE_TIME ");
                sbSQL.AppendLine("     , A.EQUIPMENT ");
                sbSQL.AppendLine("   FROM AMAT_EQPMAT_HISTORY A  ");
                sbSQL.AppendLine("   WHERE PLANT = ''CCUBEDIGITAL''  ");
                sbSQL.AppendLine("   AND TRANSACTION  = ''MEMC''  ");
                sbSQL.AppendLine("   AND EQUIPMENT LIKE ''SPU%'' ");
                sbSQL.AppendLine("   AND ITEM_CODE IN (''1B30-02008'', ''1130-00031'')  ");
                sbSQL.AppendLine("   AND INTERLOCK_ID IS NULL ");
                sbSQL.AppendLine("   AND TRANS_TIME > ''20161201'' ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" ')B ");
                sbSQL.AppendLine(" WHERE EQUIPMENT = '" + DDL_MACH.Text + "'");
                sbSQL.AppendLine(" AND NOT EXISTS (SELECT * FROM OUT_MAT_DRAIN A WHERE A.DRAIN_ID = B.MAT_LOT_NO  AND STATE_FLAG <> 'D' AND DRAIN_MAT = 'S04' AND DRAIN_MACHINE = EQUIPMENT ) ");
                sbSQL.AppendLine(" ORDER BY INPUT_TIME DESC ");

                DataTable dt = fun.getData(sbSQL.ToString());
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.NewRow();
                    dt.Rows.InsertAt(dr, 0);

                    DDL_MCSLOT_ID.DataTextField = "MAT_LOT_NO";
                    DDL_MCSLOT_ID.DataValueField = "MAT_LOT_NO";

                }
                else
                {
                    dt = new DataTable();
                    dt.Columns.Add("MAT_LOT_NO");

                    DataRow dr = dt.NewRow();
                    dt.Rows.InsertAt(dr, 0);

                }
                DDL_MCSLOT_ID.DataTextField = "MAT_LOT_NO";
                DDL_MCSLOT_ID.DataValueField = "MAT_LOT_NO";

                DDL_MCSLOT_ID.DataSource = dt;
                DDL_MCSLOT_ID.DataBind();

            }
        }

        protected void DDL_MCSLOT_ID_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDL_MAT.Text == "S01")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine(" B.DRAIN_ID ");
                sbSQL.AppendLine(" , B.ITEM_QTY ");
                sbSQL.AppendLine(" , B.LIFE_TIME ");
                sbSQL.AppendLine(" , B.DRAIN_TIME ");
                sbSQL.AppendLine(" FROM OPENQUERY(CCUBE, ' ");
                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine("     CASE WHEN EQUIPMENT LIKE ''%-%'' THEN EQUIPMENT||''-''|| TRIM(TO_CHAR(TO_NUMBER(SUBSTR(MODULE,5)), ''09''))");
                sbSQL.AppendLine("           ELSE EQUIPMENT END EQP_ID");
                sbSQL.AppendLine("      ,MAT_BATCH_ID DRAIN_ID");
                sbSQL.AppendLine("      ,ITEM_QTY ");
                sbSQL.AppendLine("      ,LIFE_TIME");
                sbSQL.AppendLine("      ,SUBSTR(TRANS_TIME, 0, 12) DRAIN_TIME");
                sbSQL.AppendLine(" FROM AMAT_EQPMAT_BATCH_STS A");
                sbSQL.AppendLine(" WHERE PLANT = ''CCUBEDIGITAL''");
                sbSQL.AppendLine(" AND STATUS = ''DRAIN'' ");
                sbSQL.AppendLine(" AND DRAIN_DATE > ''2016'' ");
                sbSQL.AppendLine(" ') B ");
                sbSQL.AppendLine(" WHERE 1=1");
                sbSQL.AppendLine(" AND B.DRAIN_ID = '" + DDL_MCSLOT_ID.Text + "'");


                DataTable dt = fun.getData(sbSQL.ToString());

                if (dt.Rows.Count > 0)
                {
                    TXT_IN_AMAT.Text = dt.Rows[0]["ITEM_QTY"].ToString();
                    TXT_IN_AMAT.Enabled = false;
                    TXT_INTGELEC.Text = dt.Rows[0]["LIFE_TIME"].ToString();
                    TXT_INTGELEC.Enabled = false;

                    TXT_DRAINDT.Text = dt.Rows[0]["DRAIN_TIME"].ToString();
                    TXT_DRAINDT.Enabled = false;
                }
            }
            else if (DDL_MAT.Text == "S03")
            {
                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine(" 	* ");
                sbSQL.AppendLine("  FROM OPENQUERY(CCUBE, ' ");
                sbSQL.AppendLine("   SELECT");
                sbSQL.AppendLine("   INPUT_TIME");
                sbSQL.AppendLine("   , LEAD(INPUT_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS OUT_TIME ");
                sbSQL.AppendLine("   , MAT_LOT_NO ");
                sbSQL.AppendLine("   , ITEM_CODE ");
                sbSQL.AppendLine("   , ITEM_QTY ");
                sbSQL.AppendLine("   , LIFE_TIME ");
                sbSQL.AppendLine("   , EQUIPMENT  ");
                sbSQL.AppendLine("   , LEAD(LIFE_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS LIFE_TIME1  ");
                sbSQL.AppendLine("   , NUM  ");
                sbSQL.AppendLine(" FROM  ");
                sbSQL.AppendLine(" (  ");
                sbSQL.AppendLine("   SELECT   ");
                sbSQL.AppendLine("     ROW_NUMBER () OVER (PARTITION BY EQUIPMENT ORDER BY TRANS_TIME) AS NUM  ");
                sbSQL.AppendLine("     , SUBSTR(A.TRANS_TIME, 0, 12) AS INPUT_TIME ");
                sbSQL.AppendLine("     , A.MAT_LOT_NO ");
                sbSQL.AppendLine("     , A.ITEM_CODE ");
                sbSQL.AppendLine("     , A.ITEM_QTY ");
                sbSQL.AppendLine("     , A.LIFE_TIME ");
                sbSQL.AppendLine("     , A.EQUIPMENT ");
                sbSQL.AppendLine("   FROM AMAT_EQPMAT_HISTORY A  ");
                sbSQL.AppendLine("   WHERE PLANT = ''CCUBEDIGITAL''  ");
                sbSQL.AppendLine("   AND TRANSACTION  = ''MEMC''  ");
                sbSQL.AppendLine("   AND EQUIPMENT LIKE ''SPU%'' ");
                sbSQL.AppendLine("   AND ITEM_CODE IN (''1B30-02008'', ''1130-00031'')  ");
                sbSQL.AppendLine("   AND INTERLOCK_ID IS NULL ");
                sbSQL.AppendLine("   AND TRANS_TIME > ''20161201'' ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" ')B ");
                sbSQL.AppendLine(" WHERE EQUIPMENT = '" + DDL_MACH.Text + "'");
                sbSQL.AppendLine(" AND MAT_LOT_NO = '" + DDL_MCSLOT_ID.Text + "'");
                sbSQL.AppendLine(" ORDER BY INPUT_TIME DESC ");

                DataTable dt = fun.getData(sbSQL.ToString());
                if (dt.Rows.Count > 0)
                {
                    TXT_IN_AMAT.Text = dt.Rows[0]["ITEM_QTY"].ToString();
                    TXT_IN_AMAT.Enabled = false;
                    TXT_INTGELEC.Text = dt.Rows[0]["LIFE_TIME1"].ToString();

                    //TXT_DRAINDT.Text = dt.Rows[0]["OUT_TIME"].ToString();

                }
              

            }
            else if (DDL_MAT.Text == "S04")
            {
                StringBuilder sbSQL = new StringBuilder();

                sbSQL.AppendLine(" SELECT ");
                sbSQL.AppendLine(" 	* ");
                sbSQL.AppendLine("  FROM OPENQUERY(CCUBE, ' ");
                sbSQL.AppendLine("   SELECT");
                sbSQL.AppendLine("   INPUT_TIME");
                sbSQL.AppendLine("   , LEAD(INPUT_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS OUT_TIME ");
                sbSQL.AppendLine("   , MAT_LOT_NO ");
                sbSQL.AppendLine("   , ITEM_CODE ");
                sbSQL.AppendLine("   , ITEM_QTY ");
                sbSQL.AppendLine("   , LIFE_TIME ");
                sbSQL.AppendLine("   , EQUIPMENT  ");
                sbSQL.AppendLine("   , LEAD(LIFE_TIME) OVER(PARTITION BY EQUIPMENT ORDER BY NUM) AS LIFE_TIME1  ");
                sbSQL.AppendLine("   , NUM  ");
                sbSQL.AppendLine(" FROM  ");
                sbSQL.AppendLine(" (  ");
                sbSQL.AppendLine("   SELECT   ");
                sbSQL.AppendLine("     ROW_NUMBER () OVER (PARTITION BY EQUIPMENT ORDER BY TRANS_TIME) AS NUM  ");
                sbSQL.AppendLine("     , SUBSTR(A.TRANS_TIME, 0, 12) AS INPUT_TIME ");
                sbSQL.AppendLine("     , A.MAT_LOT_NO ");
                sbSQL.AppendLine("     , A.ITEM_CODE ");
                sbSQL.AppendLine("     , A.ITEM_QTY ");
                sbSQL.AppendLine("     , A.LIFE_TIME ");
                sbSQL.AppendLine("     , A.EQUIPMENT ");
                sbSQL.AppendLine("   FROM AMAT_EQPMAT_HISTORY A  ");
                sbSQL.AppendLine("   WHERE PLANT = ''CCUBEDIGITAL''  ");
                sbSQL.AppendLine("   AND TRANSACTION  = ''MEMC''  ");
                sbSQL.AppendLine("   AND EQUIPMENT LIKE ''SPU%'' ");
                sbSQL.AppendLine("   AND ITEM_CODE IN (''1B30-02008'', ''1130-00031'')  ");
                sbSQL.AppendLine("   AND INTERLOCK_ID IS NULL ");
                sbSQL.AppendLine("   AND TRANS_TIME > ''20161201'' ");
                sbSQL.AppendLine(" ) ");
                sbSQL.AppendLine(" ')B ");
                sbSQL.AppendLine(" WHERE EQUIPMENT = '" + DDL_MACH.Text + "'");
                sbSQL.AppendLine(" AND MAT_LOT_NO = '" + DDL_MCSLOT_ID.Text + "'");
                sbSQL.AppendLine(" ORDER BY INPUT_TIME DESC ");

                DataTable dt = fun.getData(sbSQL.ToString());
                if (dt.Rows.Count > 0)
                {
                    TXT_IN_AMAT.Text = dt.Rows[0]["ITEM_QTY"].ToString();
                    TXT_IN_AMAT.Enabled = false;
                    TXT_INTGELEC.Text = dt.Rows[0]["LIFE_TIME1"].ToString();

                    //TXT_DRAINDT.Text = dt.Rows[0]["OUT_TIME"].ToString();

                }
                

            }
        }
    }
}
