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

namespace ERPAppAddition.ERPAddition.AM.AM_AB1002
{
    public partial class AM_AB1002 : System.Web.UI.Page
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

                DateTime setDate = DateTime.Today;
                tb_yyyymm.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00");

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
            FpSpread1.Columns.Count = 7;
            FpSpread1.Rows.Count = 0;

            FpSpread1.Columns[0].DataField = "NUM";
            FpSpread1.Columns[1].DataField = "DATE";
            FpSpread1.Columns[2].DataField = "GUBN";            
            FpSpread1.Columns[3].DataField = "AMT";
            FpSpread1.Columns[4].DataField = "ACCT";            
            FpSpread1.Columns[5].DataField = "RMK";            
            FpSpread1.Columns[6].DataField = "FLAG";            
            
            FpSpread1.Columns[0].Label = "NUM";
            FpSpread1.Columns[1].Label = "일자";
            FpSpread1.Columns[2].Label = "구분";
            FpSpread1.Columns[3].Label = "금액";
            FpSpread1.Columns[4].Label = "계정";
            FpSpread1.Columns[5].Label = "적요";
            FpSpread1.Columns[6].Label = "취소여부";

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
            string[] sl = { "","수익","비용"};
            ddlGUBN.DataSource = sl;
            ddlGUBN.DataBind();
        }


        protected void FpSpread1_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            string row = e.CommandArgument.ToString();
            //{X=2,Y=3}
            row = row.Replace("{X=", "");
            row = row.Replace("Y=", "");
            string[] arryRow = row.Split(',');
            int rr = Convert.ToInt32(arryRow[0]);            
            
            TXT_NUM.Text = FpSpread1.Sheets[0].Cells[rr, 0].Text;
            TXT_DT.Text = FpSpread1.Sheets[0].Cells[rr, 1].Text;
            ddlGUBN.Text = FpSpread1.Sheets[0].Cells[rr, 2].Text;
            TXT_AMT.Text = FpSpread1.Sheets[0].Cells[rr, 3].Text;
            TXT_ACCT.Text = FpSpread1.Sheets[0].Cells[rr,4].Text;
            TXT_RMK.Text = FpSpread1.Sheets[0].Cells[rr, 5].Text;
            TXT_FLAG.Text = FpSpread1.Sheets[0].Cells[rr, 6].Text;            

            /*선택된 row 색 칠하기*/
            for (int c = 0; c < FpSpread1.Rows.Count; c++)
            {
                FpSpread1.Sheets[0].Rows[c].BackColor = Color.Empty;
            }
            FpSpread1.Sheets[0].Rows[rr].BackColor = Color.LightPink;
        }

        protected void btn_search_click(object sender, EventArgs e)
        {
            search();
        }

        private string getSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT NUM, DT,GUBN, AMT, ACCT, REMARK, FLAG  FROM T_DAY_INCOM_NOP_INS\n");
            sbSQL.Append("WHERE DT LIKE'" + tb_yyyymm.Text + "'+'%' \n");
            sbSQL.Append("ORDER BY NUM ASC\n");
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


        protected void btn_save_click(object sender, EventArgs e)
        {
            string num = TXT_NUM.Text;
            if (num.Contains("신규추가"))
            {
                num = "";
            }
            else if (num.Equals(""))
            {
                MessageBox.ShowMessage("데이터를 선택하세요.", this.Page);
                return;
            }
            else if (TXT_FLAG.Text.Equals("Y"))
            {
                MessageBox.ShowMessage("이미 취소된 데이터는 수정이 불가합니다.", this.Page);
                return;
            }

            string dt = TXT_DT.Text;
            string gubn = ddlGUBN.Text;
            string amt = TXT_AMT.Text;
            string acct = TXT_ACCT.Text;
            string rmk = TXT_RMK.Text;
            string user = Session["User"].ToString();

            if (dt == null || dt == "" || dt.Length < 6)
                MessageBox.ShowMessage("입력일을확인하세요[yyyymmdd].", this.Page);
            else if (gubn == null || gubn == "")
                MessageBox.ShowMessage("구분을 선택하세요.", this.Page);
            else if (amt == null || amt == "")
                MessageBox.ShowMessage("금액을 입력하세요", this.Page);
            else if (acct == null || acct == "")
                MessageBox.ShowMessage("계정을 입력하세요.", this.Page);
            else
            {
                string setSQL = "";
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("MERGE  T_DAY_INCOM_NOP_INS A      \n");
                sbSQL.Append("USING (                           \n");
                sbSQL.Append("SELECT                            \n");                
                sbSQL.Append("  '" + num + "' as NUM                     \n");
                sbSQL.Append(", '" + dt + "' as DT                     \n");
                sbSQL.Append(", '" + gubn + "' as GUBN                            \n");
                sbSQL.Append(", '" + amt + "' as AMT                             \n");
                sbSQL.Append(", '" + acct + "' as ACCT                            \n");
                sbSQL.Append(", '" + rmk + "' as REMARK                          \n");
                sbSQL.Append(", '" + rmk + "' as FLAG                            \n");
                sbSQL.Append(", '" + user + "' as usr                            \n");
                sbSQL.Append(")B                                \n");
                sbSQL.Append(" ON A.NUM = B.NUM                 \n");
                sbSQL.Append("WHEN MATCHED THEN                 \n");
                sbSQL.Append("UPDATE                            \n");
                sbSQL.Append("SET                               \n");
                sbSQL.Append("    A.DT = B.DT                   \n");
                sbSQL.Append("   ,A.GUBN = B.GUBN               \n");
                sbSQL.Append("   ,A.AMT = B.AMT                 \n");
                sbSQL.Append("   ,A.ACCT = B.ACCT               \n");
                sbSQL.Append("   ,A.REMARK = B.REMARK           \n");
                sbSQL.Append("   ,A.UPDT_DT =GETDATE()          \n");
                sbSQL.Append("   ,A.UPDT_USER_ID = B.USR        \n");
                sbSQL.Append("WHEN NOT MATCHED THEN             \n");
                sbSQL.Append("INSERT (                          \n");                
                sbSQL.Append(" DT                               \n");
                sbSQL.Append(",GUBN                             \n");
                sbSQL.Append(",AMT                              \n");
                sbSQL.Append(",ACCT                             \n");
                sbSQL.Append(",REMARK                           \n");
                sbSQL.Append(",FLAG                             \n");
                sbSQL.Append(",INSRT_PROG_ID                    \n");
                sbSQL.Append(",INSRT_DT                         \n");
                sbSQL.Append(",INSRT_USER_ID                    \n");
                sbSQL.Append(",UPDT_DT                          \n");
                sbSQL.Append(",UPDT_USER_ID                     \n");
                sbSQL.Append(")VALUES(                          \n");                
                sbSQL.Append("  B.DT                            \n");
                sbSQL.Append(" ,B.GUBN                          \n");
                sbSQL.Append(" ,B.AMT                           \n");
                sbSQL.Append(" ,B.ACCT                          \n");
                sbSQL.Append(" ,B.REMARK                        \n");
                sbSQL.Append(" ,'N'                             \n");
                sbSQL.Append(" ,'AB_AB1002'                     \n");
                sbSQL.Append(" ,GETDATE()                       \n");
                sbSQL.Append(" ,B.usr                           \n");
                sbSQL.Append(" ,GETDATE()                       \n");
                sbSQL.Append(" ,B.usr                           \n");
                sbSQL.Append(" );                               \n");
                setSQL = sbSQL.ToString();

                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }
                initTextBox();
                search();
                MessageBox.ShowMessage("입력 정보가 저장 되었습니다.", this.Page);
            }
        }

        private void initTextBox()
        {
            TXT_NUM.Text = "";
            TXT_DT.Text = "";
            ddlGUBN.Text = "";
            TXT_AMT.Text = "";
            TXT_ACCT.Text = "";
            TXT_RMK.Text = "";      
        }

        protected void btn_del_click(object sender, EventArgs e)
        {
            string num = TXT_NUM.Text;
            if (num.Equals(""))
            {
                MessageBox.ShowMessage("데이터를 선택하세요.", this.Page);
                return;
            }
            else if (TXT_FLAG.Text.Equals("Y"))
            {
                MessageBox.ShowMessage("이미 취소된 데이터는 수정이 불가합니다.", this.Page);
                return;
            }
            string user = Session["User"].ToString();
            string setSQL = "";
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("UPDATE T_DAY_INCOM_NOP_INS \n");
            sbSQL.Append("SET FLAG = 'Y' \n");
            sbSQL.Append("WHERE NUM='"+num+"' \n");
            setSQL = sbSQL.ToString();

            if (QueryExecute(setSQL) < 0)
            {
                MessageBox.ShowMessage("오류 데이터가 있습니다.", this.Page);
                return;
            }
            initTextBox();
            search();
            MessageBox.ShowMessage("취소되었습니다.", this.Page);

        }

        protected void btn_ins_click(object sender, EventArgs e)
        {        
            TXT_NUM.Text = "신규추가";
            TXT_DT.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
            ddlGUBN.Text = "";
            TXT_AMT.Text = "";
            TXT_ACCT.Text = "";
            TXT_RMK.Text = "";            
        }
    }
}