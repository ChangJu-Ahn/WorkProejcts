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
using ERPAppAddition.ERPAddition.B2.B2_B0001;

namespace ERPAppAddition.ERPAddition.B2.B2_B0001
{
    public partial class B2_B0001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        int value;
        string bpcd;

        FarPoint.Web.Spread.TextCellType textCellType = new FarPoint.Web.Spread.TextCellType();
        
         
        protected void Page_Load(object sender, EventArgs e)
        {            
            if (!IsPostBack)
            {
                if (Request.QueryString["bpcd"] == null || Request.QueryString["bpcd"] == "")
                    bpcd = "BP_CD"; //erp에서 실행하지 않았을시 대비용
                else
                    bpcd = Request.QueryString["bpcd"];

                Session["bpcd"] = bpcd;

                txt_BP_CD.Text = bpcd;

                FpSpread1.Sheets[0].PageSize = 20;
                FpSpread1.Pager.Position = FarPoint.Web.Spread.PagerPosition.Top;
                FpSpread1.Pager.Mode = FarPoint.Web.Spread.PagerMode.Both;
             
                setGrid();
                FpSpread1.ActiveSheetView.AutoPostBack = true;
                FpSpread1.CommandBar.Visible = true;

                txt_MES_SHIPCD.Enabled = false;

                setPlantcdInfo();
                setSlcdInfo();

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

            FpSpread1.Sheets[0].AddColumns(0, 5);
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 0].Text = "거래처코드";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 1].Text = "MES 납품처코드";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 2].Text = "납품처명";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 3].Text = "출하공장";
            FpSpread1.Sheets[0].ColumnHeader.Cells[0, 4].Text = "출하창고";            

            FpSpread1.Sheets[0].Columns[0].Width = 100;
            FpSpread1.Sheets[0].Columns[1].Width = 100;
            FpSpread1.Sheets[0].Columns[2].Width = 150;
            FpSpread1.Sheets[0].Columns[3].Width = 100;
            FpSpread1.Sheets[0].Columns[4].Width = 100;            

            FpSpread1.Sheets[0].AddRows(0, 20);            
            //FpSpread1.Sheets[0].OperationMode = FarPoint.Web.Spread.OperationMode.ReadOnly;

            for(int i = 0; i< FpSpread1.Sheets[0].Columns.Count; i++)
            {
                FpSpread1.Sheets[0].Columns[i].HorizontalAlign = HorizontalAlign.Center;
                FpSpread1.Sheets[0].Columns[i].VerticalAlign = VerticalAlign.Middle;
                FpSpread1.Sheets[0].Columns[i].CellType = textCellType;
            }
        }
        private void setSlcdInfo()
        {
            DataTable SlcdDt = ConSql(getSlcdSQL());
            if (SlcdDt.Rows.Count > 0)
            {
                DataRow dr = SlcdDt.NewRow();
                SlcdDt.Rows.InsertAt(dr, 0);

                dnlSlcd.DataTextField = "SL_NM";
                dnlSlcd.DataValueField = "SL_CD";
                dnlSlcd.DataSource = SlcdDt;
                dnlSlcd.DataBind();
            }
        }

        private void setPlantcdInfo()
        {
            DataTable PlantDt = ConSql(getPlantcdSQL());
            if (PlantDt.Rows.Count > 0)
            {
                DataRow dr = PlantDt.NewRow();
                PlantDt.Rows.InsertAt(dr, 0);

                dnlPlantCd.DataTextField = "PLANT_NM";
                dnlPlantCd.DataValueField = "PLANT_CD";
                dnlPlantCd.DataSource = PlantDt;
                dnlPlantCd.DataBind();
            }
        }

        
        protected void Load_btn_Click(object sender, EventArgs e)
        {
            search();
            initTextBox();
        }
        private void search()
        {
            DataTable SheetDt = ConSql(getSQL());
            if (SheetDt.Rows.Count > 0)
            {
                FpSpread1.Sheets[0].DataSource = SheetDt;
                FpSpread1.DataBind();

                txt_MES_SHIPCD.Enabled = false;
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
            conn.Close();
            return resultDt;
        }

        private string getSQL()
        {
            string bpCd = txt_BP_CD.Text;
            /* 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select  						 \n");
            sbSQL.Append("	 BP_CD                       \n");
            sbSQL.Append("	,MES_SHIP_TO_PARTY_CD        \n");
            sbSQL.Append("	,MES_SHIP_TO_PARTY_NM        \n");
            sbSQL.Append("	,PLANT_CD                    \n");
            sbSQL.Append("	,SHIP_SL_CD                  \n");
            sbSQL.Append(" FROM B_BP_SHIP_TO_PARTY       \n");
            sbSQL.Append("WHERE BP_CD ='" + bpCd + "'    \n");
            sbSQL.Append("  ORDER BY MES_SHIP_TO_PARTY_CD     \n");
            return sbSQL.ToString();
        }

        private string getSlcdSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append(" SELECT ");
            sbSQL.Append("     SL_CD ");
            sbSQL.Append("    ,'[' + SL_CD + ']' + SL_NM AS SL_NM ");
            sbSQL.Append(" FROM B_STORAGE_LOCATION ");
            sbSQL.Append( "WHERE PLANT_CD = 'P04' ");
            sbSQL.Append(" ORDER BY SL_CD ");
            return sbSQL.ToString();
        }
        private string getPlantcdSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append(" SELECT ");
            sbSQL.Append("     PLANT_CD, ");
            sbSQL.Append("    '[' + PLANT_CD + ']' + PLANT_NM  AS PLANT_NM ");
            sbSQL.Append(" FROM B_PLANT ");
            sbSQL.Append(" ORDER BY PLANT_CD ");
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
            
            txt_MES_SHIPCD.Text = FpSpread1.Sheets[0].Cells[rr, 1].Text;
            txt_MES_SHIPNM.Text = FpSpread1.Sheets[0].Cells[rr, 2].Text;
            dnlPlantCd.Text = FpSpread1.Sheets[0].Cells[rr, 3].Text;
            dnlSlcd.Text = FpSpread1.Sheets[0].Cells[rr, 4].Text;
        }

        protected void dnlPlant_SelectedIndexChanged(object sender, EventArgs e)
        {
            //여기에서 다음 창고 셋팅
        }

        protected void btnNEW_Click(object sender, EventArgs e)
        {
            initTextBox();
            txt_MES_SHIPCD.Enabled = true;
        }

        private void initTextBox()
        {
            txt_MES_SHIPCD.Text = "";            
            txt_MES_SHIPNM.Text = "";
            dnlPlantCd.Text = "";
            dnlSlcd.Text = "";
        }


        protected void btnSAVE_Click(object sender, EventArgs e)
        {
            string bpCd = txt_BP_CD.Text;
            string shipCd = txt_MES_SHIPCD.Text;
            string shipNm = txt_MES_SHIPNM.Text;
            string plant = dnlPlantCd.SelectedValue;
            string slCd = dnlSlcd.SelectedValue;

            string setSQL = "";

            if (shipCd == null || shipCd == "")
                MessageBox.ShowMessage("MES 납품처 코드를 확인하세요", this.Page);
            else if (shipNm == null || shipNm == "")
                MessageBox.ShowMessage("MES 납품처 명을 확인하세요.", this.Page);

            else if (txt_MES_SHIPCD.Enabled == true)
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("insert into B_BP_SHIP_TO_PARTY \n");
                sbSQL.Append("(                      \n");
                sbSQL.Append(" BP_CD                 \n");
                sbSQL.Append(",MES_SHIP_TO_PARTY_CD  \n");
                sbSQL.Append(",MES_SHIP_TO_PARTY_NM  \n");
                sbSQL.Append(",PLANT_CD              \n");
                sbSQL.Append(",SHIP_SL_CD            \n");
                sbSQL.Append(",USER_ID               \n");
                sbSQL.Append(",ISRT_DT               \n");
                sbSQL.Append(")                      \n");
                sbSQL.Append("VALUES(                \n");
                sbSQL.Append(" '" + bpCd + "'        \n");
                sbSQL.Append(",'" + shipCd + "'      \n");
                sbSQL.Append(",'" + shipNm + "'      \n");
                sbSQL.Append(",'" + plant + "'       \n");
                sbSQL.Append(",'" + slCd + "'        \n");
                sbSQL.Append(",'B2_B0001   '         \n");
                sbSQL.Append(",GETDATE()             \n");
                sbSQL.Append(" )                     \n");
                setSQL = sbSQL.ToString();
            }
            else
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("update B_BP_SHIP_TO_PARTY \n");
                sbSQL.Append("set                      \n");
                sbSQL.Append(" MES_SHIP_TO_PARTY_NM = '" + shipNm + "' \n");
                sbSQL.Append(",PLANT_CD = '" + plant + "' \n");
                sbSQL.Append(",SHIP_SL_CD  = '" + slCd + "' \n");
                sbSQL.Append(",USER_ID  =  'USER'     \n");
                sbSQL.Append(",ISRT_DT  = GETDATE()     \n");
                sbSQL.Append("where BP_CD ='" + bpCd + "'      \n");
                sbSQL.Append("  and MES_SHIP_TO_PARTY_CD ='" + shipCd + "'      \n");
                setSQL = sbSQL.ToString();
            }

            if (QueryExecute(setSQL) < 0)
            {
                MessageBox.ShowMessage("MES납품처코드에 중복 데이터가 있습니다.", this.Page);
                return;
            }
            initTextBox();
            search();
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
            string bpCd = txt_BP_CD.Text;
            string shipCd = txt_MES_SHIPCD.Text;

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("DELETE B_BP_SHIP_TO_PARTY \n");
            sbSQL.Append("where BP_CD ='" + bpCd + "'      \n");
            sbSQL.Append("  and MES_SHIP_TO_PARTY_CD ='" + shipCd + "'      \n");

            QueryExecute(sbSQL.ToString());
            initTextBox();
            search();
            MessageBox.ShowMessage("MES 납품처코드 : [" + shipCd + "] 가 삭제되었습니다.", this.Page);
            
        }
    }
}


       