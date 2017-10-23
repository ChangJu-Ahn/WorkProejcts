using FarPoint.Web.Spread;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.MM.MM_MM003
{
    public partial class MM_MM003 : System.Web.UI.Page
    {
        SqlConnection sql_conn = new SqlConnection();
        SqlCommand sql_cmd = new SqlCommand();
        string connDBnm, userid;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {

                if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
                    connDBnm = Request.QueryString["db"].ToString();
                else
                    connDBnm = "nepes_test1";

                hdnDB.Value = connDBnm;

                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                txtBPNM.Attributes.Add("readonly", "true");
                txtITEMNM.Attributes.Add("readonly", "true");

                TimeSpan ts = new TimeSpan(-15, 0, 0, 0);
                DateTime time = DateTime.Now.Add(ts);
                //txtFrShipDT.Text = time.Year.ToString("0000") + time.Month.ToString("00") + time.Day.ToString("00");
                //txtToShipDT.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");


                InitiGrid();


            }

            sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings[hdnDB.Value].ConnectionString);
            sql_cmd = new SqlCommand();
        }

        private void InitiGrid()
        {

            FarPoint.Web.Spread.CheckBoxCellType chkboxCelltype = new FarPoint.Web.Spread.CheckBoxCellType();
            SheetView sv = FpSpread1.Sheets[0];
            //FarPoint.Web.Spread.

            sv.RowCount = 0;
            sv.ColumnCount = 0;
            sv.DataSource = null;

            FarPoint.Web.Spread.Extender.DateCalendarCellType dtcalCell = new FarPoint.Web.Spread.Extender.DateCalendarCellType();
            dtcalCell.DateFormat = "yyyy-MM-dd";

            FarPoint.Web.Spread.CheckBoxCellType chkCell = new CheckBoxCellType();
            FarPoint.Web.Spread.IntegerCellType intCell = new IntegerCellType();
            FarPoint.Web.Spread.DoubleCellType doubleCell = new DoubleCellType();
            FarPoint.Web.Spread.ComboBoxCellType cmbCell = new ComboBoxCellType();
                        
            System.Globalization.NumberFormatInfo nfd = new NumberFormatInfo();

            nfd.CurrencyDecimalSeparator = ",";

            doubleCell.NumberFormat = nfd;

            //doubleCell.DecimalDigits = 2;

            System.Globalization.NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalDigits = 0;
            nfi.CurrencySymbol = "";

            nfi.GetFormat(System.Type.GetType("CurrencyCellType"));


            intCell.NumberFormat = nfi;

            intCell.NumberFormat.CurrencyDecimalSeparator = ",";

            chkCell.AutoPostBack = true;


            cls_FpSheet sheet = new cls_FpSheet();

            sheet.AddColumnHeader(sv, "CHK", " ", 20, HorizontalAlign.Center, false, 0, chkCell);
            sheet.AddColumnHeader(sv, "PLANT_CD", "공장", 80, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "PUR_GRP", "구매그룹", 80, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "PUR_GRP_NM", "구매그룹명", 100, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "PUR_ORG", "제품명", 80, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "PR_NO", "PR 번호", 100, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "PR_STS", "PR 상태", 70, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "SPPL_CD", "공급처코드", 80, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "BP_NM", "공급처명", 60, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "ITEM_CD", "품목코드", 80, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "ITEM_NM", "품목명", 270, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "ITEM_NM_DESC", "품목상세정보", 180, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "DLVY_DT", "필요일", 120, HorizontalAlign.Center, true, 0, dtcalCell);
            sheet.AddColumnHeader(sv, "REQ_DT", "요청일", 120, HorizontalAlign.Center, true, 0, dtcalCell);
            sheet.AddColumnHeader(sv, "REQ_PRSN", "요청자", 70, HorizontalAlign.Center, true);
            sheet.AddColumnHeader(sv, "REQ_QTY", "요청량", 80, HorizontalAlign.Right, true, 0, doubleCell);
            sheet.AddColumnHeader(sv, "REQ_UNIT", "단위", 70, HorizontalAlign.Left, true);
            sheet.AddColumnHeader(sv, "ORD_QTY", "발주량", 80, HorizontalAlign.Right, true, 0, doubleCell);
            sheet.AddColumnHeader(sv, "RES_QTY", "요청잔량", 80, HorizontalAlign.Right, true, 0, doubleCell);
            sheet.AddColumnHeader(sv, "RCPT_QTY", "입고량", 80, HorizontalAlign.Right, true, 0, doubleCell);
            sheet.AddColumnHeader(sv, "IV_QTY", "매입량", 80, HorizontalAlign.Right, true, 0, doubleCell);

            //int sheetSize = 0;
            //for(int i=0; i<FpSpread1.Columns.Count; i++)
            //{
            //    sheetSize += FpSpread1.Columns[i].Width;
            //}

        }

        protected void btnExport_Click(object sender, EventArgs e)
        {

        }

        protected void bntSearch_Click(object sender, EventArgs e)
        {
            GetSearch();
        }

        private void GetSearch()
        {

            StringBuilder sb = new StringBuilder();


            InitiGrid();

            sb.AppendLine("");
            sb.AppendLine("SELECT ");
            sb.AppendLine("  0 AS CHK");
            sb.AppendLine("	, AA.PLANT_CD ");
            sb.AppendLine("	, SP.PUR_GRP ");
            sb.AppendLine("	, PG.PUR_GRP_NM ");
            sb.AppendLine("	, AA.PUR_ORG ");
            sb.AppendLine("	, AA.PR_NO ");
            sb.AppendLine("	, PR_STS ");
            sb.AppendLine("	, ISNULL(SP.SPPL_CD, AA.SPPL_CD) AS SPPL_CD ");
            sb.AppendLine("	, ( SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD = isnull(SP.SPPL_CD, AA.SPPL_CD) ) AS BP_NM ");
            sb.AppendLine("	, AA.ITEM_CD ");
            sb.AppendLine("	, IT.ITEM_NM ");
            sb.AppendLine("	, ISNULL(BB.ITEM_NM, '') AS ITEM_NM_DESC ");
            sb.AppendLine("	, AA.DLVY_DT ");
            sb.AppendLine("	, AA.REQ_DT ");
            sb.AppendLine("	, AA.REQ_PRSN ");
            sb.AppendLine("	, AA.REQ_QTY ");
            sb.AppendLine("	, AA.REQ_UNIT ");
            sb.AppendLine("	, AA.ORD_QTY ");
            sb.AppendLine("	, AA.REQ_QTY - AA.ORD_QTY AS RES_QTY ");
            sb.AppendLine("	, AA.RCPT_QTY ");
            sb.AppendLine("	, AA.IV_QTY ");
            sb.AppendLine(" FROM M_PUR_REQ AA WITH(NOLOCK)  ");
            sb.AppendLine("	 LEFT OUTER JOIN S_SO_TRACKING BB  WITH(NOLOCK) ");
            sb.AppendLine("		ON AA.TRACKING_NO = BB.TRACKING_NO  ");
            sb.AppendLine("	 LEFT OUTER JOIN M_PR_QUOTA_BY_SPPL SP WITH(NOLOCK) ");
            sb.AppendLine("		ON AA.PR_NO = SP.PR_NO ");
            sb.AppendLine("	 INNER JOIN B_ITEM IT ");
            sb.AppendLine("		ON AA.ITEM_CD = IT.ITEM_CD ");
            sb.AppendLine("	 INNER JOIN B_PUR_GRP PG ");
            sb.AppendLine("	    ON SP.PUR_GRP = PG.PUR_GRP ");
            sb.AppendLine(" WHERE 1=1 --AA.PR_NO = '' ");
            sb.AppendLine("  AND ISNULL(AA.CLS_FLG, '') <> 'Y' ");
            sb.AppendLine("  --AND (AA.REQ_QTY - AA.ORD_QTY) <> 0 ");
            sb.AppendLine("  AND AA.REQ_QTY > AA.ORD_QTY ");
            sb.AppendLine("  AND AA.PR_STS NOT IN ('RQ', 'CF') ");

            if (txtBPCD.Text.Length > 0)
            {
                sb.AppendLine("  AND ISNULL(SP.SPPL_CD, AA.SPPL_CD) = '" + txtBPCD.Text + "' ");
            }

            if (txtITEMCD.Text.Length > 0)
            {
                sb.AppendLine("  AND AA.ITEM_CD = '" + txtITEMCD.Text + "' ");
            }

            if (txtPUR_GRP.Text.Length > 0)
            {
                sb.AppendLine("  AND SP.PUR_GRP = '" + txtPUR_GRP.Text + "' ");
            }

            if (txtPLANT_CD.Text.Length > 0)
            {
                sb.AppendLine("  AND AA.PLANT_CD = '" + txtPLANT_CD.Text + "' ");
            }
            if (txtFrREQ_DT.Text.Length > 0 && txtToREQ_DT.Text.Length > 0 )
            {
                sb.AppendLine("  AND AA.REQ_DT BETWEEN '" + txtFrREQ_DT.Text + "' AND '" + txtToREQ_DT + "'");
            }

            if (txtFrDLVY_DT.Text.Length > 0 && txtToDLVY_DT.Text.Length > 0)
            {
                sb.AppendLine("  AND AA.DLVY_DT BETWEEN '" + txtFrDLVY_DT.Text + "' AND '" + txtToDLVY_DT + "'");
            }

            sb.AppendLine("ORDER BY AA.PLANT_CD, ISNULL(SP.SPPL_CD, AA.SPPL_CD), SP.PUR_GRP, AA.DLVY_DT ");




          

            DataTable dt = getData(sb.ToString());

            if (dt.Rows.Count > 0)
            {

                FarPoint.Web.Spread.Model.DefaultSheetDataModel model = new FarPoint.Web.Spread.Model.DefaultSheetDataModel(dt);

                FpSpread1.Sheets[0].DataModel = model;
                FpSpread1.DataBind();
            }
            return;

        }

        public DataTable getData(string sql)
        {
            DataSet ds = new DataSet();

            DataTable retDt = new DataTable();

            try
            {
                if (sql_conn.Database == "")
                {
                    sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings[hdnDB.Value].ConnectionString);
                    sql_cmd = new SqlCommand();
                }
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = sql;

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds);
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                retDt = ds.Tables[0];

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
            return retDt;
        }
        protected void btnSave_Click(object sender, EventArgs e)
        {

        }
    }
}