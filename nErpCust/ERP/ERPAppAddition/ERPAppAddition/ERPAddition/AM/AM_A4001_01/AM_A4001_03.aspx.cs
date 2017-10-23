using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.AM.AM_A4001_03;

namespace ERPAppAddition.ERPAddition.AM.AM_A4001_03
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_enc"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        int value;

        FarPoint.Web.Spread.TextCellType textCellType = new FarPoint.Web.Spread.TextCellType();


        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                FpSpread1.Sheets[0].PageSize = 20;
                FpSpread1.Pager.Position = FarPoint.Web.Spread.PagerPosition.Top;
                FpSpread1.Pager.Mode = FarPoint.Web.Spread.PagerMode.Both;


                setGrid();
                FpSpread1.ActiveSheetView.AutoPostBack = true;
                FpSpread1.CommandBar.Visible = true;

                setBankInfo();
                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void setGrid()
        {

            int colIndex = FpSpread1.Sheets[0].ColumnCount;
            FpSpread1.Sheets[0].RemoveColumns(0, FpSpread1.Sheets[0].ColumnCount);
            FpSpread1.Sheets[0].RemoveRows(0, FpSpread1.Sheets[0].Rows.Count);

            FpSpread1.Sheets[0].AddColumns(0, 6);
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 0].Text = "사번";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 1].Text = "이름";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 2].Text = "은행코드";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 3].Text = "은행명";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 4].Text = "계좌번호";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 5].Text = "수정시간";

            FpSpread1.Sheets[0].Columns[0].Width = 80;
            FpSpread1.Sheets[0].Columns[1].Width = 80;
            FpSpread1.Sheets[0].Columns[2].Width = 70;
            FpSpread1.Sheets[0].Columns[3].Width = 200;
            FpSpread1.Sheets[0].Columns[4].Width = 200;
            FpSpread1.Sheets[0].Columns[5].Width = 80;

            FpSpread1.Sheets[0].AddRows(0, 20);
            //FpSpread1.Sheets[0].OperationMode = FarPoint.Web.Spread.OperationMode.ReadOnly;

            for (int i = 0; i < FpSpread1.Sheets[0].Columns.Count; i++)
            {
                FpSpread1.Sheets[0].Columns[i].HorizontalAlign = HorizontalAlign.Center;
                FpSpread1.Sheets[0].Columns[i].VerticalAlign = VerticalAlign.Middle;
                FpSpread1.Sheets[0].Columns[i].CellType = textCellType;
            }

        }
        private void setBankInfo()
        {
            DataTable BankDt = ConSql(getBankSQL());
            if (BankDt.Rows.Count > 0)
            {
                DataRow dr = BankDt.NewRow();
                BankDt.Rows.InsertAt(dr, 0);

                dnlBANKNM.DataTextField = "MINOR_NM";
                dnlBANKNM.DataValueField = "MINOR_CD";
                dnlBANKNM.DataSource = BankDt;
                dnlBANKNM.DataBind();

            }
        }

        protected void Load_btn_Click(object sender, EventArgs e)
        {
            search();
        }
        private void search()
        {
            txt_EMPNO.Enabled = false;
            DataTable SheetDt = ConSql(getSQL());
            if (SheetDt.Rows.Count > 0)
            {
                FpSpread1.Sheets[0].DataSource = SheetDt;
                FpSpread1.DataBind();
            }
            else
            {
                FpSpread1.ActiveSheetView.Rows.Remove(0, FpSpread1.ActiveSheetView.Rows.Count);
                MessageBox.ShowMessage("조회된 내용이 없습니다.", this.Page);
            }
        }

        private DataTable ConSql(string SQL)
        {
            DataTable resultDt = new DataTable();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = SQL;

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(resultDt);
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

            return resultDt;
        }

        private string getSQL()
        {
            string EMPNO = txtFEMP_NO.Text;
            string NAME = txtFNAME.Text;
            /* 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select  						\n");
            sbSQL.Append("	EMP_NO                      \n");
            sbSQL.Append("	,NAME                       \n");
            sbSQL.Append("	,BANK                       \n");
            sbSQL.Append("	,BN.MINOR_NM AS BANK_NM     \n");
            sbSQL.Append("	,BANK_ACCNT                 \n");
            sbSQL.Append("	,HR.UPDT_DT                 \n");
            sbSQL.Append("from HDF020T_HR HR            \n");
            sbSQL.Append("   ,B_MINOR BN                \n");
            sbSQL.Append("WHERE HR.BANK = BN.MINOR_CD   \n");
            sbSQL.Append("  AND BN.MAJOR_CD='ZZ008'     \n");
            if (EMPNO != "") sbSQL.Append("  AND HR.EMP_NO='" + EMPNO + "'  \n");
            if (NAME != "") sbSQL.Append("  AND HR.NAME='" + NAME + "'     \n");
            sbSQL.Append("  ORDER BY EMP_NO     \n");
            return sbSQL.ToString();
        }

        private string getBankSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append(" Select MINOR_CD,MINOR_NM ");
            sbSQL.Append(" From   B_MINOR ");
            sbSQL.Append("Where  MAJOR_CD='ZZ008' order by MINOR_NM");
            return sbSQL.ToString();
        }


        protected void FpSpread1_CellClick(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            string row = e.CommandArgument.ToString();
            //{X=2,Y=3}
            row = row.Replace("{X=", "");
            row = row.Replace("Y=", "");
            string[] arryRow = row.Split(',');
            int rr = Convert.ToInt32(arryRow[0]);

            txt_EMPNO.Text = FpSpread1.Sheets[0].Cells[rr, 0].Text;
            txt_NAME.Text = FpSpread1.Sheets[0].Cells[rr, 1].Text;
            txt_BANKCD.Text = FpSpread1.Sheets[0].Cells[rr, 2].Text;
            dnlBANKNM.Text = FpSpread1.Sheets[0].Cells[rr, 2].Text;
            txt_BANKADD.Text = FpSpread1.Sheets[0].Cells[rr, 4].Text;
            txt_UPDT.Text = FpSpread1.Sheets[0].Cells[rr, 5].Text;
        }

        protected void dnlBANKCD_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_BANKCD.Text = dnlBANKNM.SelectedValue.ToString();
        }

        protected void btnNEW_Click(object sender, EventArgs e)
        {
            initTextBox();
        }

        private void initTextBox()
        {
            txt_EMPNO.Text = "";
            txt_NAME.Text = "";
            txt_BANKCD.Text = "";
            dnlBANKNM.Text = "";
            txt_BANKADD.Text = "";
            txt_UPDT.Text = "";
            txt_EMPNO.Enabled = true;
        }


        protected void btnSAVE_Click(object sender, EventArgs e)
        {
            //추가(2014.03.02)

            string EMP_NO = txt_EMPNO.Text;
            string NAME = txt_NAME.Text;
            string BANK = txt_BANKCD.Text;
            string BANK_ACCNT = txt_BANKADD.Text;

            if (EMP_NO == null || EMP_NO == "")
                MessageBox.ShowMessage("사번을 입력하세요.", this.Page);
            else if (NAME == null || NAME == "")
                MessageBox.ShowMessage("이름을 입력하세요.", this.Page);
            else if (BANK == null || BANK == "")
                MessageBox.ShowMessage("은행명을 확인하세요", this.Page);
            else if (BANK_ACCNT == null || BANK_ACCNT == "")
                MessageBox.ShowMessage("계좌번호를 입력해주세요.", this.Page);

            else
            {
                string setSQL = "";
                if (txt_UPDT == null || txt_UPDT.Text == "")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("insert into HDF020T_HR \n");
                    sbSQL.Append("(                      \n");
                    sbSQL.Append(" EMP_NO                \n");
                    sbSQL.Append(",NAME                  \n");
                    sbSQL.Append(",BANK                  \n");
                    sbSQL.Append(",BANK_ACCNT            \n");
                    sbSQL.Append(",ISRT_DT               \n");
                    sbSQL.Append(",ISRT_EMP_NO           \n");
                    sbSQL.Append(",UPDT_DT               \n");
                    sbSQL.Append(",UPDT_EMP_NO           \n");
                    sbSQL.Append(")                      \n");
                    sbSQL.Append("VALUES(                \n");
                    sbSQL.Append("'" + EMP_NO + "'       \n");
                    sbSQL.Append(",'" + NAME + "'         \n");
                    sbSQL.Append(",'" + BANK + "'         \n"); ;
                    sbSQL.Append(",'" + BANK_ACCNT + "'   \n");
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(",'AM_A4001_01'          \n");
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(",'AM_A4001_01'         \n");
                    sbSQL.Append(" )                     \n");
                    setSQL = sbSQL.ToString();
                }
                else
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("update HDF020T_HR \n");
                    sbSQL.Append("set                      \n");
                    sbSQL.Append(" NAME = '" + NAME + "' \n");
                    sbSQL.Append(",BANK = '" + BANK + "' \n");
                    sbSQL.Append(",BANK_ACCNT  = '" + BANK_ACCNT + "' \n");
                    sbSQL.Append(",UPDT_DT  = GETDATE()               \n");
                    sbSQL.Append(",UPDT_EMP_NO  = 'AM_A4001_01_u'     \n");
                    sbSQL.Append("where EMP_NO ='" + EMP_NO + "'      \n");
                    setSQL = sbSQL.ToString();
                }

                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("사번 : [" + EMP_NO + "] 중복 데이터가 있습니다.", this.Page);
                    return;
                }
                initTextBox();
                search();
                MessageBox.ShowMessage("사번 : [" + EMP_NO + "] 정보가 저장 되었습니다.", this.Page);
            }
        }

        public int QueryExecute(string sql)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                value = -1;
            }
            conn.Close();
            return value;
        }

        protected void btnDEL_Click(object sender, EventArgs e)
        {
            string EMP_NO = txt_EMPNO.Text;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("DELETE HDF020T_HR \n");
            sbSQL.Append("where EMP_NO ='" + EMP_NO + "'      \n");

            QueryExecute(sbSQL.ToString());
            initTextBox();
            search();
            MessageBox.ShowMessage("사번 : [" + EMP_NO + "] 정보가 삭제되었습니다.", this.Page);

        }
    }
}


