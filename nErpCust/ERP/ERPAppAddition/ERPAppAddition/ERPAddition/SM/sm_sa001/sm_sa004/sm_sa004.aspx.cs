using FarPoint.Web.Spread;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.SM.sm_sa001.sm_sa004
{
    public partial class sm_sa004 : System.Web.UI.Page
    {
         sa_fun fun = new sa_fun();

        //string strConn = ConfigurationManager.AppSettings["connectionKey"];
        //string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();

        DataTable _dtDate = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string userid;
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                initiSpread(FpSpread1.Sheets[0]);

                TimeSpan ts = new TimeSpan(-180, 0, 0, 0);

                DateTime date = DateTime.Now.Date.Add(ts);
//                cal_fr_yyyymm.SelectedDate = DateTime.Now.Date.Add(ts);
//                cal_to_yyyymm.SelectedDate = DateTime.Now.Date;

                tb_fr_yyyymm.Text = date.Year.ToString("0000") + date.Month.ToString("00");
                tb_to_yyyymm.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00");

                SetComboBoxScrap();
            }

        }

        private void initiSpread( SheetView sv)
        {
            DataTable dt = new DataTable();

            FarPoint.Web.Spread.Model.DefaultSheetDataModel model = new FarPoint.Web.Spread.Model.DefaultSheetDataModel(dt);
            //model.AutoGenerateColumns = false;
            sv.DataModel = model;

            FpSpread1.DataBind();
            
        }


        public void AddColumnHeader(FarPoint.Web.Spread.SheetView sv, string sDataFild, string sText, int Width, HorizontalAlign hAlign, bool Visible = true, bool Lock = false, int row = 0, bool merge = false)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();

            if (sv.DataSource == null)
            {
                dt = new DataTable();
                ds.Tables.Add(dt);
                //sv.AutoPostBack = true;

                //FpSpread1.EnableClientScript = false;
                sv.OperationMode = FarPoint.Web.Spread.OperationMode.Normal;

                FarPoint.Web.Spread.Model.DefaultSheetDataModel model = new FarPoint.Web.Spread.Model.DefaultSheetDataModel(dt);
                //model.AutoGenerateColumns = false;
                sv.DataModel = model;
                sv.DataMember = "Table1";
                sv.DataSourceID = "Table1";

                sv.AllowPage = true;
                sv.PageSize = 25;




                FpSpread1.DataBind();
            }
            else
            {
                if(sv.DataSource.GetType() == typeof(DataSet))
                {
                    ds = (DataSet)sv.DataSource;
                     dt = ds.Tables[0];
                }
                else
                {
                    dt = (DataTable)sv.DataSource;
                    // dt = ds.Tables[0];
                }
                

            }
            

            dt.Columns.Add(sDataFild);

            dt.AcceptChanges();

            //sv.DataSource = dt;


            FarPoint.Web.Spread.Model.DefaultSheetDataModel modeladd = new FarPoint.Web.Spread.Model.DefaultSheetDataModel(dt);

            //modeladd.DataSource = dt;
            sv.DataModel = modeladd;
            FpSpread1.DataBind();

            modeladd.AutoGenerateColumns = false;


            int col = sv.ColumnHeader.Columns.Count - 1;
            sv.ColumnHeader.Cells[row, col].Text = sText;
            sv.ColumnHeader.Columns[col].DataField = sDataFild;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sv.Columns[i].DataField = dt.Columns[i].ColumnName;
            }
            sv.Columns[col].Locked = Lock;

            sv.Columns[col].HorizontalAlign = hAlign;

            sv.ColumnHeader.Columns[col].Width = Width;
            sv.ColumnHeader.Columns[col].Visible = Visible;

            if (merge)
            {
                sv.SetColumnMerge(col, FarPoint.Web.Spread.Model.MergePolicy.Always);
            }
        }

        private void SetComboBoxScrap()
        {
            /* */
            //DataTable UNIT = fun.getData("SELECT GROUP1_DESC, GROUP2_CODE   FROM DBO.SA_SYS_CODE WHERE GROUP1_CODE IN ('S01', 'S02', 'S03') GROUP BY GROUP1_DESC, GROUP2_CODE");
            //if (UNIT.Rows.Count > 0)
            //{
            //    cmbScrap.DataTextField = "GROUP1_DESC";
            //    cmbScrap.DataValueField = "GROUP2_CODE";
            //    cmbScrap.DataSource = UNIT;
            //    cmbScrap.DataBind();
            //}
            string sSql = "SELECT GROUP4_DESC AS ITEM_NM, GROUP2_CODE AS ITEM_CD   FROM DBO.SA_SYS_CODE WHERE GROUP1_CODE IN ('S01', 'S02', 'S03') GROUP BY GROUP4_DESC, GROUP2_CODE";
            SetcmbVelue(sSql, cmbScrap);
        }


        protected void btnSearch_Click(object sender, EventArgs e)
        {
            if(cmbScrap.SelectedIndex == 0)
            {
                string script = "alert(\"구분을 선택해주세요..\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                return;
            }
            if (cmbPlant.SelectedIndex == 0)
            {
                string script = "alert(\"공장을 선택해주세요..\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                return;
            }
            if (tb_fr_yyyymm.Text.Length == 6 && tb_to_yyyymm.Text.Length == 6)
            {
                _dtDate = GetDataDt();

                if(_dtDate.Rows.Count < 1)
                {
                    string script = "alert(\"조회된 Data가 없습니다.\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                string sSQL = "";
                string sMap = "";
                string sScrap = "";
                FpSpread1.ActiveSheetViewIndex = 0;

                if (cmbScrap.SelectedValue.ToString() == "ETCH")
                {
                    sSQL = GetEtchSQL();
                    sMap = "rd_sm_sa004.rdlc";
                    sScrap = "ETCH";
                }
                else if (cmbScrap.SelectedValue.ToString() == "PLAT")
                {
                    sSQL = GetPlatSQL();
                    sMap = "rd_sm_sa004_plat.rdlc";
                    sScrap = "PGC";
                }
                else if (cmbScrap.SelectedValue.ToString() == "SPUTTER")
                {
                    sSQL = GetSputterSQL();
                    sMap = "rd_sm_sa004_target.rdlc";
                    sScrap = "Target";
                }

                DataSet ds = GetDataSet(sSQL);

                SetSheet(ds.Tables["DataSet11"], sScrap);
                SetSheetDtl(ds.Tables["DataSet12"], sScrap);
                if (sScrap == "PGC")
                {
                    DataSet dsPgc = EQPDataSet(ds.Tables["DataSet1"]);
                    DrawChart(dsPgc, sMap);
                }
                else
                {
                    DrawChart(ds.Tables["DataSet1"], sMap);
                }

                
                
            }
        }

        private DataSet EQPDataSet(DataTable dt)
        {
            DataSet ds = new DataSet();
            DataView dv = new DataView(dt);
            DataTable dtEQPList = dv.ToTable(true, "MACHINE");

            for (int i = 0; i < dtEQPList.Rows.Count; i++ )
            {

                DataTable dtAdd = dt.Clone();

                dtAdd = dt.Select("MACHINE = '"+dtEQPList.Rows[i]["MACHINE"].ToString()+"'").CopyToDataTable();
                //if(dr.Length > 0)
                //{
                //    for(int)
                //    dtAdd.ImportRow(dr)
                //}

                dtAdd.TableName = "DataTable"+i;

                ds.Tables.Add(dtAdd);
            }

            if(ds.Tables.Count < 4)
            {
                int addDtCnt = 4 - ds.Tables.Count;
                for (int i = 0; i < addDtCnt; i++)
                {
                    DataTable dtAdd = dt.Clone();
                    dtAdd.TableName = "DataTable" + (ds.Tables.Count+i);
                    ds.Tables.Add(dtAdd);
                }
            }


                return ds;
        }

        private void DrawChart(DataTable dt, string smap)
        {
            ReportViewer1.Reset();
            ReportViewer1.LocalReport.ReportPath = Server.MapPath(smap);
            ReportViewer1.LocalReport.DisplayName = "_Au 회수 Reporting_" + DateTime.Now.ToShortDateString();
            ReportDataSource rds = new ReportDataSource();
            rds.Name = "DataSet1";
            rds.Value = dt;
            ReportViewer1.LocalReport.DataSources.Add(rds);
        }
        private void DrawChart(DataSet ds, string smap)
        {
            ReportViewer1.Reset();
            ReportViewer1.LocalReport.ReportPath = Server.MapPath(smap);
            ReportViewer1.LocalReport.DisplayName = "_Au 회수 Reporting_" + DateTime.Now.ToShortDateString();

            
            
            for (int i = 0; i < ds.Tables.Count; i++)
            {

                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet"+(i+1);
                rds.Value = ds.Tables[i];
                //rds.DataSourceId = "Chart1" + (i+1);
                //ReportViewer1.LocalReport.
                rds.DataMember = "Chart1" + (i + 1);
                ReportViewer1.LocalReport.DataSources.Add(rds);
                
            }
        }
        private void SetSheet(DataTable dt, string sFlag)
        {
            SheetView sv = FpSpread1.Sheets[0];

            initiSpread(sv);

            
            
            AddColumnHeader(sv, "MACHINE", "장비", 80, HorizontalAlign.Center, true, true,0, true);

            int gubun = 1;
            if (sFlag.Equals("Target"))
            {
                AddColumnHeader(sv, "DRAIN_MAT", "Scrap 종류", 100, HorizontalAlign.Center, true, true, 0, true);
                gubun = 2;
            }
            AddColumnHeader(sv, "GUBUN", "구분", 110, HorizontalAlign.Center, true, true);

            if (_dtDate.Rows.Count > 0)
            {
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    AddColumnHeader(sv, _dtDate.Rows[i]["DT"].ToString(), _dtDate.Rows[i]["DT"].ToString(), 80, HorizontalAlign.Right, true, true);
                }
            }
            AddColumnHeader(sv, "AVG_QTY", "평균", 80, HorizontalAlign.Right, true, true);

            sv.DataSource = dt;
            
            for(int i=0; i< sv.Rows.Count; i++)
            {
                if (sv.Cells[i, gubun].Text.Contains("회수율"))
                {
                    PercentCellType celltype = new PercentCellType();
                    celltype.DecimalDigits = 2;
                    sv.Rows[i].CellType = celltype;
                }
                else
                {
                    System.Globalization.NumberFormatInfo nfi = new NumberFormatInfo();
                    nfi.NumberDecimalDigits = 2;
                    nfi.CurrencySymbol = "";
                    nfi.GetFormat(System.Type.GetType("CurrencyCellType"));

                    DoubleCellType celltype = new FarPoint.Web.Spread.DoubleCellType();

                    celltype.NumberFormat = nfi;
                    //celltype.FormatString = "999,999,999";
                    sv.Rows[i].CellType = celltype;
                }
            }

            sv.SheetName = "Trend";
        }

        private void SetSheetDtl(DataTable dt, string sFlag)
        {
            SheetView sv = FpSpread1.Sheets[1];

            initiSpread(sv);
            
            AddColumnHeader(sv, "PLANT", "공장", 40, HorizontalAlign.Center, true, true, 0, true);
            AddColumnHeader(sv, "DRAIN_DT", "일자", 70, HorizontalAlign.Center, true, true, 0, true);
            AddColumnHeader(sv, "MACHINE", "장비", 60, HorizontalAlign.Center, true, true, 0, true);
            if (!sFlag.Equals("ETCH"))
            {
                AddColumnHeader(sv, "DRAIN_ID", "Scrap ID", 120, HorizontalAlign.Center, true, true, 0, false);
            }
            AddColumnHeader(sv, "DRAIN_PROCESS", "구분", 50, HorizontalAlign.Center, true, true, 0, false);
            AddColumnHeader(sv, "DRAIN_MAT", "Scrap 종류", 50, HorizontalAlign.Center, true, true, 0, false);

            if (sFlag.Equals("ETCH"))
            {
                AddColumnHeader(sv, "DRAIN_IN_DT", "Scrap 발생일", 70, HorizontalAlign.Center, true, true, 0, false);
            }   

            if (!sFlag.Equals("ETCH"))
            {
                AddColumnHeader(sv, "AMAT_IN", "투입량", 50, HorizontalAlign.Center, true, true, 0, false);
            }          
            if (sFlag.Equals("Target"))
            {
                AddColumnHeader(sv, "DRAIN_INTGELEC", "전산전력", 50, HorizontalAlign.Center, true, true, 0, false);
            
            }
            else if(sFlag.Equals("PGC"))
            {
                AddColumnHeader(sv, "DRAIN_INTGELEC", "전산전류", 50, HorizontalAlign.Center, true, true, 0, false);
            }
            if (!sFlag.Equals("ETCH"))
            {
                AddColumnHeader(sv, "AMAT_USG", "사용량", 50, HorizontalAlign.Center, true, true, 0, false);
            }
            AddColumnHeader(sv, "DRAIN_QTY", "누적장수", 50, HorizontalAlign.Center, true, true, 0, false);

            AddColumnHeader(sv, "DRAIN_SCRAP_QTY", "Scrap 수량", 50, HorizontalAlign.Center, true, true, 0, false);
            AddColumnHeader(sv, "DRAIN_UINT", "단위", 50, HorizontalAlign.Center, true, true, 0, false);

            AddColumnHeader(sv, "R_QTY", "반출량", 50, HorizontalAlign.Center, true, true, 0, false);
            AddColumnHeader(sv, "R_QTY_UNIT", "단위", 40, HorizontalAlign.Center, true, true, 0, false);

            AddColumnHeader(sv, "R_DOC_NO", "반출문서", 150, HorizontalAlign.Center, true, true, 0, false);
            AddColumnHeader(sv, "R_DT", "반출일", 70, HorizontalAlign.Center, true, true, 0, false);

            if (!sFlag.Equals("Target"))
            {
                AddColumnHeader(sv, "IN_AU_QTY", "농도", 50, HorizontalAlign.Center, true, true, 0, false);
                AddColumnHeader(sv, "IN_AU_UNIT", "단위", 40, HorizontalAlign.Center, true, true, 0, false);
            }
            AddColumnHeader(sv, "IN_QTY", "Au 회수", 50, HorizontalAlign.Center, true, false);
            AddColumnHeader(sv, "IN_QTY_UNIT", "Au 회수", 50, HorizontalAlign.Center, true, false);

            if (sFlag.Equals("PGC"))
            {
                AddColumnHeader(sv, "IN_QTY_PGC", "PGC 회수", 50, HorizontalAlign.Center, true, true, 0, false);
            }
            


            sv.DataSource = dt;

            //for (int i = 0; i < sv.Rows.Count; i++)
            //{
            //    if (sv.Cells[i, gubun].Text.Contains("회수율"))
            //    {
            //        PercentCellType celltype = new PercentCellType();
            //        celltype.DecimalDigits = 2;
            //        sv.Rows[i].CellType = celltype;
            //    }
            //    else
            //    {
            //        System.Globalization.NumberFormatInfo nfi = new NumberFormatInfo();
            //        nfi.NumberDecimalDigits = 2;
            //        nfi.CurrencySymbol = "";
            //        nfi.GetFormat(System.Type.GetType("CurrencyCellType"));

            //        DoubleCellType celltype = new FarPoint.Web.Spread.DoubleCellType();

            //        celltype.NumberFormat = nfi;
            //        //celltype.FormatString = "999,999,999";
            //        sv.Rows[i].CellType = celltype;
            //    }
            //}
            sv.SheetName = "상세";
        }

        private DataSet GetDataSet(string sSQL)
        {
            DataSet ds = new DataSet();

            try
            {
                // 프로시져 실행: 기본데이타 생성
               

                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = sSQL;

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
                    //DataRow dr = ds.Tables["DataSet1"].NewRow();
                    //ds.Tables["DataSet1"].Rows.InsertAt(dr, 0);
                    //FpSpread1.DataSource = ds.Tables["DataSet1"];
                    //FpSpread1.DataBind();

                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                }

                //FpSpread1.DataSource = ds.Tables["DataSet1"];
                //FpSpread1.DataBind();

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }

            return ds;
        }

        private DataTable GetDataDt()
        {
            DataTable dt = new DataTable();

            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                //sql_cmd.CommandText = "SELECT DISTINCT SUBSTRING(CONVERT(CHAR(10),DATEADD(D,NUMBER,'"+tb_fr_yyyymm.Text+"01'),112), 0, 7) DT ";
                //sql_cmd.CommandText += " FROM MASTER..SPT_VALUES ";
                //sql_cmd.CommandText += " WHERE TYPE = 'P' AND NUMBER <= DATEDIFF(D,'" + tb_fr_yyyymm.Text + "01','" + tb_to_yyyymm.Text + "01') ";

                StringBuilder sSQL = new StringBuilder();

                sSQL.AppendLine(" SELECT");
                sSQL.AppendLine(" SUBSTRING(DRAIN_IN_DT, 0, 7) AS DT ");
                sSQL.AppendLine(" FROM OUT_MAT_DRAIN");
                sSQL.AppendLine(" WHERE 1=1");
                sSQL.AppendLine(" AND  STATE_FLAG <> 'D'");
                //sSQL.AppendLine(" AND  AMAT_IN IS NOT NULL");
                sSQL.AppendLine(" AND DRAIN_IN_DT BETWEEN '" + tb_fr_yyyymm.Text + "01' AND '" + tb_to_yyyymm.Text + "31'");
                if (cmbScrap.SelectedIndex > 0)
                {
                    sSQL.AppendLine(" AND DRAIN_PROCESS = '" + cmbScrap.SelectedValue.ToString() + "'");
                }
                if (cmbEqpID.SelectedIndex > 0)
                {
                    sSQL.AppendLine(" AND DRAIN_MACHINE = '" + cmbEqpID.SelectedValue.ToString() + "'");
                }
                if (cmbPlant.SelectedIndex > 0)
                {
                    sSQL.AppendLine(" AND DRAIN_PLANT = '" + cmbPlant.SelectedValue.ToString() + "'");
                }
                sSQL.AppendLine(" GROUP BY SUBSTRING(DRAIN_IN_DT, 0, 7)");

                sql_cmd.CommandText = sSQL.ToString();

                DataSet ds = new DataSet();
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
                    
                }


                dt = ds.Tables["DataSet1"];
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }

            return dt;
        }

        private string GetPlatSQL()
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sbM = new StringBuilder();
            StringBuilder sbDtl = new StringBuilder();

            sbDtl.AppendLine(" SELECT  ");
            sbDtl.AppendLine("  		 BB.DRAIN_PLANT AS PLANT ");
            sbDtl.AppendLine("  		, SUBSTRING(BB.DRAIN_IN_DT, 0, 7) AS DRAIN_DT ");
            sbDtl.AppendLine("  		, BB.DRAIN_MACHINE AS MACHINE ");
            sbDtl.AppendLine("  		, BB.DRAIN_ID ");           
            sbDtl.AppendLine("  		, BB.DRAIN_PROCESS ");
            sbDtl.AppendLine("  		, (SELECT DISTINCT GROUP1_DESC FROM SA_SYS_CODE WHERE GROUP1_CODE = BB.DRAIN_MAT AND PLANT_CD = BB.DRAIN_PLANT) AS DRAIN_MAT ");
            sbDtl.AppendLine("  		, AMAT_IN AS AMAT_IN  -- 투입량 ");
            sbDtl.AppendLine("  		, CASE WHEN DRAIN_INTGELEC_MCS IS NULL THEN DRAIN_INTGELEC ELSE DRAIN_INTGELEC_MCS END AS DRAIN_INTGELEC ");
            sbDtl.AppendLine("  		, CASE WHEN ISNULL(AA.R_MATE_QTY,0) <> 0 THEN ROUND(CONVERT(float, CASE WHEN DRAIN_INTGELEC_MCS IS NULL THEN DRAIN_INTGELEC ELSE DRAIN_INTGELEC_MCS END) * CONVERT(float,UD_MINOR_NM)/ (SNED_OUTQTY / AA.R_MATE_QTY),2) END AS AMAT_USG ");
            //sbDtl.AppendLine("  		, ROUND(CONVERT(float, CASE WHEN DRAIN_INTGELEC_MCS IS NULL THEN DRAIN_INTGELEC ELSE DRAIN_INTGELEC_MCS END) * CONVERT(float,UD_MINOR_NM), 2) AS AMAT_USG ");
            sbDtl.AppendLine("  		, DRAIN_QTY AS DRAIN_QTY --누적장수 ");
            sbDtl.AppendLine("  		, DRAIN_SCRAP_QTY AS DRAIN_SCRAP_QTY --SCRAP수량 ");
            sbDtl.AppendLine("  		, DRAIN_UINT AS DRAIN_UINT --SCRAP단위 ");
            sbDtl.AppendLine("  	    , AA.R_QTY AS R_QTY --반출수량 ");
            sbDtl.AppendLine("  		, AA.R_QTY_UNIT AS R_QTY_UNIT --반출단위 ");
            sbDtl.AppendLine("  		, AA.R_DOC_NO --반출문서 ");
            sbDtl.AppendLine("  		, R_DT --반출일 ");
            
            sbDtl.AppendLine("  		, IN_AU_QTY AS IN_AU_QTY --Au농도 ");
            sbDtl.AppendLine("  		, IN_AU_UNIT AS IN_AU_UNIT --농도단위 ");
            sbDtl.AppendLine("  		, IN_QTY AS IN_QTY --Au회수수량 ");

            sbDtl.AppendLine("  		, R_MATE_QTY AS IN_QTY_PGC  --Au회수수량 ");
            //sbDtl.AppendLine("  		, CASE WHEN ISNULL(AA.R_MATE_QTY,0) <> 0 THEN IN_QTY / (SNED_OUTQTY / AA.R_MATE_QTY) END AS IN_QTY_PGC  --Au회수수량 ");
            sbDtl.AppendLine("  		, IN_QTY_UNIT AS IN_QTY_UNIT --Au회수단위 ");
            //sbDtl.AppendLine("  		, SEND_DT AS SEND_DT --사급자재지급일 ");
            //sbDtl.AppendLine("  		, SEND_QTY AS SEND_QTY --사급자재지급량 ");
            //sbDtl.AppendLine("  		, SNED_OUTQTY AS SNED_OUTQTY --임가공량 ");
            sbDtl.AppendLine("  	FROM OUT_MAT_HIS AA WITH(NOLOCK) ");
            sbDtl.AppendLine("  	    INNER JOIN OUT_MAT_DRAIN BB WITH(NOLOCK) ");
            sbDtl.AppendLine("  	    ON AA.LOT_NO = BB.LOT_NO ");
            sbDtl.AppendLine("  	    INNER JOIN B_USER_DEFINED_MINOR CC WITH(NOLOCK) ");
            sbDtl.AppendLine("  	    ON BB.DRAIN_PLANT = CC.UD_MINOR_CD ");
            sbDtl.AppendLine(" 		AND CC.UD_MAJOR_CD = 'SA001' ");
            sbDtl.AppendLine("  	WHERE  1=1 ");
            sbDtl.AppendLine("  	 AND BB.DRAIN_PROCESS = 'PLAT'");
            sbDtl.AppendLine("  	 AND BB.STATE_FLAG <> 'D' ");
            sbDtl.AppendLine("  	 AND BB.AMAT_IN IS NOT NULL ");
            if (cmbPlant.SelectedIndex > 0)
            {
                sbDtl.AppendLine("  AND BB.DRAIN_PLANT = '" + cmbPlant.SelectedValue.ToString() + "'");
            }
            if (cmbEqpID.SelectedIndex > 0)
            {
                sbDtl.AppendLine(" 	 AND BB.DRAIN_MACHINE = '" + cmbEqpID.SelectedValue.ToString() + "'");
            }

            sbDtl.AppendLine("   AND DRAIN_IN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "' ");

           
            sbM.AppendLine(" SELECT ");
            sbM.AppendLine("  DRAIN_DT ");
            sbM.AppendLine("  , MACHINE ");
            sbM.AppendLine("  , PLANT ");
            sbM.AppendLine("  , ROUND(MAX(AMAT_IN), 2) AS AMAT_IN ");
            sbM.AppendLine("  , MAX(DRAIN_INTGELEC) AS DRAIN_INTGELEC  ");
            sbM.AppendLine("  , MAX(AMAT_USG) AS AMAT_USG  ");
            sbM.AppendLine("  , ROUND(MAX(DRAIN_SCRAP_QTY), 2) AS DRAIN_SCRAP_QTY ");
            sbM.AppendLine("  , ROUND(SUM(R_QTY), 2) AS R_QTY ");
            sbM.AppendLine("  , ROUND(SUM(IN_QTY), 2) AS IN_AU_QTY ");
            sbM.AppendLine("  , ROUND(AVG(IN_AU_QTY), 2) AS AU_M ");
            sbM.AppendLine("  , ROUND(SUM(IN_QTY_PGC), 2) AS IN_QTY_PGC ");

            sbM.AppendLine("  , CASE WHEN ISNULL(SUM(IN_QTY_PGC),0) <> 0 THEN (ROUND(SUM(IN_QTY_PGC), 2)) / ROUND(MAX(AMAT_IN), 2) END AS IN_PGC_RATION  ");
            //sbM.AppendLine("  , (ROUND(SUM(IN_QTY_PGC), 2)) / ROUND(MAX(AMAT_IN), 2) AS IN_PGC_RATION  ");

            sbM.AppendLine("  , ROUND(((SUM(IN_QTY)/ 1000) / SUM(R_QTY)) * 100, 2) AS IN_AU_RATION ");
            sbM.AppendLine("  , ROUND((MAX(AMAT_USG) + ROUND(SUM(IN_QTY_PGC), 2) - ROUND(MAX(AMAT_IN), 2)),2) AS AMAT_LOSS ");

            sbM.AppendLine(" FROM ");
            sbM.AppendLine(" (  ");

            sbM.AppendLine(sbDtl.ToString());

            sbM.AppendLine("   )A  ");
            sbM.AppendLine(" WHERE 1=1 ");
            sbM.AppendLine(" GROUP BY DRAIN_DT, MACHINE , PLANT ");
          

            sb.AppendLine(" SELECT ");
            sb.AppendLine(" DRAIN_DT ");

            sb.AppendLine(" ,MACHINE "); //Test

            sb.AppendLine(" , SUM(AMAT_IN) AS AMAT_IN  ");
            sb.AppendLine(" , SUM(AMAT_USG) AS AMAT_USG  ");
            sb.AppendLine(" , SUM(CONVERT(float,DRAIN_INTGELEC)) AS DRAIN_INTGELEC ");
            sb.AppendLine(" , SUM(R_QTY) AS R_QTY ");
            sb.AppendLine(" , SUM(IN_AU_QTY) AS IN_AU_QTY ");
            sb.AppendLine(" , SUM(IN_QTY_PGC) AS IN_QTY_PGC ");
            sb.AppendLine(" , AVG(AU_M) AS AU_M ");
            sb.AppendLine(" , AVG(IN_PGC_RATION) AS IN_PGC_RATION ");
            sb.AppendLine(" , AVG(IN_AU_RATION) AS IN_AU_RATION ");
            sb.AppendLine(" , SUM(AMAT_LOSS) AS AMAT_LOSS ");
            sb.AppendLine(" FROM( ");
            sb.AppendLine(sbM.ToString());
            sb.AppendLine(" )AA ");
            //sb.AppendLine(" GROUP BY DRAIN_DT ");

            sb.AppendLine(" GROUP BY DRAIN_DT, MACHINE ");

            sb.AppendLine("; ");

            sb.AppendLine(" WITH  ");
            sb.AppendLine(" DRAIN AS ");
            sb.AppendLine(" ( ");
            sb.AppendLine(sbM.ToString());
            sb.AppendLine(" ) ");

            sb.AppendLine(" SELECT ");
            sb.AppendLine("  MACHINE ");
            sb.AppendLine("  , GUBUN ");
            

            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , MAX(a" + _dtDate.Rows[i]["DT"].ToString() + ") AS '" + _dtDate.Rows[i]["DT"].ToString() + "'");
            }
            sb.AppendLine("  , MAX(AVG_QTY) AS AVG_QTY ");
            sb.AppendLine("  , SEQ ");
            sb.AppendLine(" FROM ");
            sb.AppendLine(" ( ");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '투입량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_IN END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_IN)) AS AVG_QTY ");
            sb.AppendLine("   , 0 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '사용량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_USG END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_USG)) AS AVG_QTY ");
            sb.AppendLine("   , 1 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");

            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Scrap 량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_SCRAP_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,DRAIN_SCRAP_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 2 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");

            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '회수량'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_QTY_PGC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_QTY_PGC)) AS AVG_QTY ");
            sb.AppendLine("   , 3 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Loss량'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_LOSS END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_LOSS)) AS AVG_QTY ");
            sb.AppendLine("   , 4 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '회수율'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_PGC_RATION END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_PGC_RATION)) AS AVG_QTY ");
            sb.AppendLine("  , 5 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '적산전류'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , MAX(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_INTGELEC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,DRAIN_INTGELEC)) AS AVG_QTY ");
            sb.AppendLine("   , 6 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '반출량(Kg)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 7 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Au 농도'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AU_M END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AU_M)) AS AVG_QTY ");
            sb.AppendLine("   , 8 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            if (false)//cmbEqpID.SelectedIndex < 1)
            {

                sb.AppendLine("   UNION ALL ( ");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '투입량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_IN END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_IN)) AS AVG_QTY ");
                sb.AppendLine("   , 0 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '사용량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_USG END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_USG)) AS AVG_QTY ");
                sb.AppendLine("   , 1 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '회수량'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_QTY_PGC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_QTY_PGC)) AS AVG_QTY ");
                sb.AppendLine("   , 2 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '차액(Loss)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_LOSS END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_LOSS)) AS AVG_QTY ");
                sb.AppendLine("   , 3 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'PGC 회수율'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , AVG(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_PGC_RATION END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_PGC_RATION)) AS AVG_QTY ");
                sb.AppendLine("  , 4 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '적산전류'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , MAX(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_INTGELEC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,DRAIN_INTGELEC)) AS AVG_QTY ");
                sb.AppendLine("   , 5 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '반출량(Kg)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 5 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'Au 농도'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AU_M END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AU_M)) AS AVG_QTY ");
                sb.AppendLine("   , 6 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("  GROUP BY MACHINE ");
                sb.AppendLine("   ) ");
            }
            sb.AppendLine(" ) A");
            sb.AppendLine("   GROUP BY MACHINE, GUBUN, SEQ ");
            sb.AppendLine("   ORDER BY MACHINE, SEQ ");

            sb.AppendLine("; ");
            sbDtl.AppendLine(" ORDER BY PLANT, MACHINE, DRAIN_DT  ");

            sb.AppendLine(sbDtl.ToString());

            return sb.ToString();
        }

        private string GetSputterSQL()
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sbM = new StringBuilder();
            StringBuilder sbDtl = new StringBuilder();

            sbDtl.AppendLine("  SELECT  ");
            sbDtl.AppendLine(" 		SUBSTRING(BB.DRAIN_IN_DT, 0, 7) AS DRAIN_DT ");
            sbDtl.AppendLine(" 		, BB.DRAIN_ID ");
            sbDtl.AppendLine(" 		, BB.DRAIN_MACHINE AS MACHINE ");
            sbDtl.AppendLine(" 		, BB.DRAIN_PLANT AS PLANT ");
            sbDtl.AppendLine(" 		, BB.DRAIN_PROCESS ");
            sbDtl.AppendLine(" 		, CASE WHEN BB.DRAIN_MAT = 'S04' THEN 'SHIELD' ELSE 'AU TARGET' END AS DRAIN_MAT ");
            sbDtl.AppendLine(" 		, CASE WHEN BB.DRAIN_MAT = 'S04' THEN NULL ELSE (AMAT_IN) END AS AMAT_IN  -- 투입량 ");
            sbDtl.AppendLine(" 		, CASE WHEN BB.DRAIN_MAT = 'S04' THEN NULL ELSE CASE WHEN (DRAIN_INTGELEC_MCS) IS NULL THEN (DRAIN_INTGELEC) ELSE (DRAIN_INTGELEC_MCS) END END AS DRAIN_INTGELEC ");
            sbDtl.AppendLine("      , ROUND(CONVERT(float, CASE WHEN BB.DRAIN_MAT = 'S04' THEN NULL ELSE CASE WHEN (DRAIN_INTGELEC_MCS) IS NULL THEN (DRAIN_INTGELEC) ELSE (DRAIN_INTGELEC_MCS) END END) * CONVERT(float,UD_MINOR_NM),2) AS AMAT_USG  ");
            sbDtl.AppendLine(" 		, (DRAIN_QTY) AS DRAIN_QTY --누적장수 ");
            sbDtl.AppendLine(" 		, (CASE WHEN DRAIN_UINT = 'KG' THEN DRAIN_SCRAP_QTY * 1000 ELSE DRAIN_SCRAP_QTY END ) AS DRAIN_SCRAP_QTY --SCRAP수량 ");
            sbDtl.AppendLine(" 		, 'GR' AS DRAIN_UINT ");
            sbDtl.AppendLine(" 	    , (CASE WHEN R_QTY_UNIT = 'KG' THEN AA.R_QTY * 1000 ELSE AA.R_QTY END) AS R_QTY --반출수량 ");
            sbDtl.AppendLine(" 		, 'GR' AS R_QTY_UNIT --반출단위 ");
            sbDtl.AppendLine(" 		, AA.R_DOC_NO --반출문서 ");
            sbDtl.AppendLine("      , R_DT --반출일  ");
            sbDtl.AppendLine(" 		, (IN_QTY) AS IN_QTY --Au회수수량 ");
            sbDtl.AppendLine(" 		, (IN_QTY_UNIT) AS IN_QTY_UNIT --Au회수단위 ");
            //sbDtl.AppendLine(" 		, (SEND_DT) AS SEND_DT --사급자재지급일 ");
            //sbDtl.AppendLine(" 		, (SEND_QTY) AS SEND_QTY --사급자재지급량 ");
            //sbDtl.AppendLine(" 		, (SNED_OUTQTY) AS SNED_OUTQTY --임가공량 ");
            sbDtl.AppendLine(" 	FROM OUT_MAT_HIS AA WITH(NOLOCK) ");
            sbDtl.AppendLine(" 	   INNER JOIN OUT_MAT_DRAIN BB WITH(NOLOCK) ");
            sbDtl.AppendLine(" 	     ON AA.LOT_NO = BB.LOT_NO ");
            sbDtl.AppendLine(" 	   INNER JOIN B_USER_DEFINED_MINOR CC WITH(NOLOCK) ");
            sbDtl.AppendLine(" 	     ON BB.DRAIN_PLANT = CC.UD_MINOR_CD  ");
            sbDtl.AppendLine(" 	     AND CC.UD_MAJOR_CD = 'SA002'  ");
            sbDtl.AppendLine(" 	WHERE  1=1 ");
            sbDtl.AppendLine(" 	 AND BB.DRAIN_PROCESS = 'SPUTTER'");
            sbDtl.AppendLine(" 	 AND BB.STATE_FLAG <> 'D' ");
            if (cmbPlant.SelectedIndex > 0)
            {
                sbDtl.AppendLine(" AND BB.DRAIN_PLANT = '" + cmbPlant.SelectedValue.ToString() + "'");
            }
            if (cmbEqpID.SelectedIndex > 0)
            {
                sbDtl.AppendLine(" 	 AND BB.DRAIN_MACHINE = '" + cmbEqpID.SelectedValue.ToString() + "'");
            }
            sbDtl.AppendLine("   AND  DRAIN_IN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");


         
            sbM.AppendLine(" SELECT ");
            sbM.AppendLine("  DRAIN_DT ");
            sbM.AppendLine("  , MACHINE ");
            sbM.AppendLine("  , PLANT ");
            sbM.AppendLine("  , DRAIN_MAT ");
            sbM.AppendLine("  , ROUND(SUM(AMAT_IN), 2) AS AMAT_IN ");
            sbM.AppendLine("  , SUM(CONVERT(float,DRAIN_INTGELEC)) AS DRAIN_INTGELEC ");
            sbM.AppendLine("  , SUM(AMAT_USG) AS AMAT_USG ");
            sbM.AppendLine("  , ROUND(SUM(R_QTY), 2) AS R_QTY ");
            sbM.AppendLine("  , ROUND(SUM(DRAIN_SCRAP_QTY), 2) AS DRAIN_SCRAP_QTY  ");
            sbM.AppendLine("  , ROUND(SUM(IN_QTY), 2) AS IN_AU_QTY ");
            sbM.AppendLine("  , CASE WHEN SUM(AMAT_IN) > 0 THEN ROUND(SUM(IN_QTY) / SUM(AMAT_IN) ,2) ELSE NULL END  AS IN_AU_RATION ");
            sbM.AppendLine("  , ROUND(((SUM(IN_QTY)) / SUM(R_QTY)), 4) AS IN_SD_AU_RATION ");
            sbM.AppendLine("  , SUM(AMAT_USG) + SUM(IN_QTY) - ROUND(SUM(AMAT_IN), 2) AS AMAT_LOSS  ");
            sbM.AppendLine(" FROM ");
            sbM.AppendLine(" (  ");

            sbM.AppendLine(sbDtl.ToString());
            sbM.AppendLine(" ) A  ");

            sbM.AppendLine(" GROUP BY DRAIN_DT, MACHINE, PLANT, DRAIN_MAT ");
            

            sb.AppendLine(" SELECT ");
            sb.AppendLine(" DRAIN_DT ");
            sb.AppendLine(" , SUM(AMAT_IN) AS AMAT_IN  ");
            sb.AppendLine(" , SUM(AMAT_USG) AS AMAT_USG  ");
            sb.AppendLine(" , SUM(CONVERT(float,DRAIN_INTGELEC)) AS DRAIN_INTGELEC ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'AU TARGET' THEN DRAIN_SCRAP_QTY END) AS DRAIN_SCRAP_QTY  ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'AU TARGET' THEN R_QTY END)  AS R_QTY ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'SHIELD' THEN R_QTY END) AS R_SHD_QTY  ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'SHIELD' THEN DRAIN_SCRAP_QTY END) AS DRAIN_SCRAP_SHD_QTY ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_QTY END) AS IN_AU_QTY ");
            sb.AppendLine(" , SUM(CASE WHEN DRAIN_MAT = 'SHIELD' THEN IN_AU_QTY END) AS IN_AU_SHD_QTY ");
            sb.AppendLine(" , CASE WHEN SUM(AMAT_IN) > 0 THEN ROUND(SUM(CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_QTY END) / SUM(AMAT_IN) ,2) ELSE NULL END  AS IN_AU_RATION  ");
            sb.AppendLine(" , ROUND((SUM(CASE WHEN DRAIN_MAT = 'SHIELD' THEN IN_AU_QTY END) / SUM(CASE WHEN DRAIN_MAT = 'SHIELD' THEN R_QTY END)), 2) AS IN_AU_SHD_RATION ");
            sb.AppendLine(" , ROUND(SUM(AMAT_LOSS), 2) AS AMAT_LOSS  ");
            sb.AppendLine(" FROM( ");
            sb.AppendLine(sbM.ToString());
            sb.AppendLine(" )AA ");
            sb.AppendLine(" GROUP BY DRAIN_DT ");

            sb.AppendLine("; ");

            sb.AppendLine(" WITH  ");
            sb.AppendLine(" DRAIN AS ");
            sb.AppendLine(" ( ");
            sb.AppendLine(sbM.ToString());
            sb.AppendLine(" ) ");

            sb.AppendLine(" SELECT ");
            sb.AppendLine("  MACHINE ");
            sb.AppendLine("  , DRAIN_MAT ");
            sb.AppendLine("  , GUBUN ");


            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , MAX(a" + _dtDate.Rows[i]["DT"].ToString() + ") AS '" + _dtDate.Rows[i]["DT"].ToString() + "'");
            }
            sb.AppendLine("  , MAX(AVG_QTY) AS AVG_QTY ");
            sb.AppendLine("  , SEQ ");
            sb.AppendLine(" FROM ");
            sb.AppendLine(" ( ");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '투입량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_IN END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_IN)) AS AVG_QTY ");
            sb.AppendLine("   , 0 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT  ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '사용량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_USG END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_USG)) AS AVG_QTY ");
            sb.AppendLine("   , 1 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , 'Scrap량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_SCRAP_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,DRAIN_SCRAP_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 2 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '회수량'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_AU_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 2 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , 'Loss량'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_LOSS END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,AMAT_LOSS)) AS AVG_QTY ");
            sb.AppendLine("   , 4 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");

            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '회수율'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_RATION ELSE IN_SD_AU_RATION END END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_RATION ELSE IN_SD_AU_RATION END)) AS AVG_QTY ");
            sb.AppendLine("  , 5 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '적산전력'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , MAX(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_INTGELEC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,DRAIN_INTGELEC)) AS AVG_QTY ");
            sb.AppendLine("   , 6 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , DRAIN_MAT ");
            sb.AppendLine("   , '반출량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 3 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE, DRAIN_MAT ");
            
            if (false)//cmbEqpID.SelectedIndex < 1)
            {

                sb.AppendLine("   UNION ALL ( ");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("     , DRAIN_MAT");
                sb.AppendLine("   , '투입량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_IN END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_IN)) AS AVG_QTY ");
                sb.AppendLine("   , 0 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , '사용량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_USG END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_USG)) AS AVG_QTY ");
                sb.AppendLine("   , 1 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , '회수량'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_AU_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 2 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , '차액(Loss)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN AMAT_LOSS END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,AMAT_LOSS)) AS AVG_QTY ");
                sb.AppendLine("   , 4 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , 'Au 회수율'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , AVG(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_RATION ELSE IN_SD_AU_RATION END END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,CASE WHEN DRAIN_MAT = 'AU TARGET' THEN IN_AU_RATION ELSE IN_SD_AU_RATION END)) AS AVG_QTY ");
                sb.AppendLine("  , 5 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , '적산전력'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , MAX(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_INTGELEC END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,DRAIN_INTGELEC)) AS AVG_QTY ");
                sb.AppendLine("   , 6 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("     AND  DRAIN_MAT <> 'SHIELD' ");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , DRAIN_MAT ");
                sb.AppendLine("   , '반출량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 3 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine("  GROUP BY DRAIN_MAT ");
                sb.AppendLine("   ) ");
            }
            sb.AppendLine(" ) A");
            sb.AppendLine("   GROUP BY MACHINE, DRAIN_MAT, GUBUN, SEQ ");
            sb.AppendLine("   ORDER BY MACHINE, DRAIN_MAT, SEQ ");

            sb.AppendLine(" ; ");

            sbDtl.AppendLine(" ORDER BY MACHINE, DRAIN_DT, DRAIN_MAT");

            sb.AppendLine(sbDtl.ToString());


            return sb.ToString();
        }


        private string GetEtchSQL()
        {
            StringBuilder sbM = new StringBuilder();
            StringBuilder sb = new StringBuilder();

            StringBuilder sbDtl = new StringBuilder();

            sbDtl.AppendLine("  SELECT  ");
            sbDtl.AppendLine(" 		SUBSTRING(R_DT, 0, 7) AS DRAIN_DT ");
            sbDtl.AppendLine(" 		, BB.DRAIN_IN_DT AS DRAIN_IN_DT ");
            sbDtl.AppendLine(" 		, BB.DRAIN_PLANT AS PLANT ");
            sbDtl.AppendLine(" 		, BB.DRAIN_MACHINE AS MACHINE ");
            sbDtl.AppendLine(" 		, DRAIN_PROCESS ");
            sbDtl.AppendLine(" 		, DRAIN_MAT ");
            sbDtl.AppendLine(" 		, AA.LOT_NO ");
            sbDtl.AppendLine(" 		, (DRAIN_QTY) AS DRAIN_QTY --누적장수 ");
            sbDtl.AppendLine(" 		, (DRAIN_SCRAP_QTY) AS DRAIN_SCRAP_QTY --SCRAP수량 ");
            sbDtl.AppendLine(" 		, (DRAIN_UINT) AS DRAIN_UINT --SCRAP단위 ");
            sbDtl.AppendLine(" 	    , (AA.R_QTY) AS R_QTY --반출수량 ");
            sbDtl.AppendLine(" 		, (AA.R_QTY_UNIT) AS R_QTY_UNIT --반출단위 ");
            sbDtl.AppendLine(" 		, AA.R_DOC_NO ");
            sbDtl.AppendLine(" 		, AA.R_DT ");
            sbDtl.AppendLine(" 		, (IN_AU_QTY) AS IN_AU_QTY --Au농도 ");
            sbDtl.AppendLine(" 		, (IN_AU_UNIT) AS IN_AU_UNIT --농도단위 ");
            sbDtl.AppendLine(" 		, (IN_QTY) AS IN_QTY --Au회수수량 ");
            sbDtl.AppendLine(" 		, (IN_QTY_UNIT) AS IN_QTY_UNIT --Au회수단위 ");
            //sbDtl.AppendLine(" 		, (SEND_DT) AS SEND_DT --사급자재지급일 ");
            //sbDtl.AppendLine(" 		, (SEND_QTY) AS SEND_QTY --사급자재지급량 ");
            //sbDtl.AppendLine(" 		, (SNED_OUTQTY) AS SNED_OUTQTY --임가공량 ");
            sbDtl.AppendLine(" 	FROM OUT_MAT_HIS AA WITH(NOLOCK) ");
            sbDtl.AppendLine(" 	   INNER JOIN  OUT_MAT_DRAIN BB WITH(NOLOCK) ");
            sbDtl.AppendLine(" 	   ON  AA.LOT_NO = BB.LOT_NO ");
            sbDtl.AppendLine(" 	WHERE 1=1 ");
            sbDtl.AppendLine(" 	 AND BB.DRAIN_PROCESS = '" + cmbScrap.SelectedValue.ToString() + "'");
            sbDtl.AppendLine(" 	 AND BB.STATE_FLAG <> 'D'");
            if (cmbEqpID.SelectedIndex > 0)
            {
                sbDtl.AppendLine(" 	 AND BB.DRAIN_MACHINE = '" + cmbEqpID.SelectedValue.ToString() + "'");
            }
            if (cmbPlant.SelectedIndex > 0)
            {
                sbDtl.AppendLine(" AND BB.DRAIN_PLANT = '" + cmbPlant.SelectedValue.ToString() + "'");
            }
            sbDtl.AppendLine("   AND  R_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");

            //
            sbM.AppendLine(" SELECT ");
            sbM.AppendLine("  DRAIN_DT "); ;
            sbM.AppendLine("  , SUM(DRAIN_SCRAP_QTY) AS DRAIN_SCRAP_QTY ");
            sbM.AppendLine("  , SUM(R_QTY) AS R_QTY ");
            sbM.AppendLine("  , SUM(IN_QTY) AS IN_QTY ");
            sbM.AppendLine("  , AVG(IN_AU_QTY) AS IN_AU_QTY ");
            //sbM.AppendLine("  , AVG(IN_AU_RATION) AS IN_AU_RATION ");
            sbM.AppendLine("  , ROUND(((SUM(IN_QTY)/ 1000) / SUM(R_QTY)), 4)  AS IN_AU_RATION ");
            //
            sbM.AppendLine(" FROM (");

            sbM.AppendLine(" SELECT ");
            sbM.AppendLine("  DRAIN_DT ");
            sbM.AppendLine("  , MACHINE ");
            sbM.AppendLine("  , LOT_NO ");
            sbM.AppendLine("  , ROUND(SUM(DRAIN_SCRAP_QTY), 2) AS DRAIN_SCRAP_QTY ");
            sbM.AppendLine("  , ROUND(MAX(R_QTY), 2)  AS R_QTY   ");
            sbM.AppendLine("  , ROUND(MAX(IN_QTY), 2) AS IN_QTY ");
            sbM.AppendLine("  , ROUND(AVG(IN_AU_QTY), 2) AS IN_AU_QTY ");
            sbM.AppendLine("  , ROUND(((MAX(IN_QTY)/ 1000) / SUM(R_QTY)), 4) AS IN_AU_RATION ");
            sbM.AppendLine(" FROM ");
            //sbM.AppendLine(" ( SELECT   ");
            //sbM.AppendLine("    DRAIN_PLANT   ");
            //sbM.AppendLine("  , MACHINE   ");
            //sbM.AppendLine("  , DRAIN_DT   ");
            //sbM.AppendLine("  , SUM(DRAIN_QTY) AS DRAIN_QTY   ");
            //sbM.AppendLine("  , SUM(DRAIN_SCRAP_QTY) AS DRAIN_SCRAP_QTY   ");
            //sbM.AppendLine("  , MAX(R_QTY) AS R_QTY   ");
            //sbM.AppendLine("  , DRAIN_PLANT   ");
            sbM.AppendLine(" (   ");

            sbM.AppendLine(sbDtl.ToString());

            sbM.AppendLine(" ) A   ");

            sbM.AppendLine(" GROUP BY DRAIN_DT, MACHINE, PLANT, LOT_NO");
            //
            sbM.AppendLine(" ) AA");
            sbM.AppendLine(" GROUP BY DRAIN_DT ");

            sb.AppendLine(sbM.ToString());

           


            sb.AppendLine(" ; ");
            sb.AppendLine(" WITH  ");
            sb.AppendLine(" DRAIN AS ");
            sb.AppendLine(" ( ");
            sb.AppendLine(" SELECT ");
            sb.AppendLine("  DRAIN_DT ");
            sb.AppendLine("  , MACHINE ");
            sb.AppendLine("  , SUM(DRAIN_SCRAP_QTY) AS DRAIN_SCRAP_QTY ");
            sb.AppendLine("  , SUM(R_QTY) AS R_QTY ");
            sb.AppendLine("  , SUM(IN_QTY) AS IN_QTY ");
            sb.AppendLine("  , AVG(IN_AU_QTY) AS IN_AU_QTY ");
            sb.AppendLine("  , ROUND(((SUM(IN_QTY)/ 1000) / SUM(R_QTY)), 4)  AS IN_AU_RATION ");
            sb.AppendLine(" FROM (");
            sb.AppendLine(" SELECT ");
            sb.AppendLine("  DRAIN_DT ");
            sb.AppendLine("  , MACHINE ");
            sb.AppendLine("  , LOT_NO ");
            sb.AppendLine("  , ROUND(SUM(DRAIN_SCRAP_QTY), 2) AS DRAIN_SCRAP_QTY ");
            sb.AppendLine("  , ROUND(MAX(R_QTY), 2)  AS R_QTY   ");
            sb.AppendLine("  , ROUND(MAX(IN_QTY), 2) AS IN_QTY ");
            sb.AppendLine("  , ROUND(AVG(IN_AU_QTY), 2) AS IN_AU_QTY ");

            sb.AppendLine(" FROM ");
            sb.AppendLine(" ( ");
            sb.AppendLine(sbDtl.ToString());
            sb.AppendLine(" )A ");
            sb.AppendLine(" GROUP BY DRAIN_DT, MACHINE, PLANT, LOT_NO");
            //
            sb.AppendLine(" ) AA");
            sb.AppendLine(" GROUP BY DRAIN_DT, MACHINE ");

            sb.AppendLine(" ) ");

            sb.AppendLine(" SELECT ");
            sb.AppendLine("  MACHINE ");
            sb.AppendLine("  , GUBUN ");


            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , MAX(a" + _dtDate.Rows[i]["DT"].ToString() + ") AS '" + _dtDate.Rows[i]["DT"].ToString() + "'");
            }
            sb.AppendLine(" , MAX(AVG_QTY) AS AVG_QTY");
            sb.AppendLine("   ,SEQ ");
            sb.AppendLine(" FROM ");
            sb.AppendLine(" ( ");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Scrap량(Kg)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_SCRAP_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,DRAIN_SCRAP_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 0 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , '반출량(Kg)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 1 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Au 농도'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , AVG(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_AU_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 2 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Au 회수량(g)'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_QTY)) AS AVG_QTY ");
            sb.AppendLine("   , 3 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");
            sb.AppendLine(" UNION ALL");
            sb.AppendLine("   SELECT ");
            sb.AppendLine("     MACHINE ");
            sb.AppendLine("   , 'Au 회수율'  AS GUBUN ");
            for (int i = 0; i < _dtDate.Rows.Count; i++)
            {
                sb.AppendLine(" , AVG(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_RATION END) AS a" + _dtDate.Rows[i]["DT"].ToString());
            }
            sb.AppendLine("   , AVG(CONVERT(float,IN_AU_RATION)) AS AVG_QTY ");
            sb.AppendLine("   , 4 AS SEQ ");
            sb.AppendLine("   FROM DRAIN ");
            sb.AppendLine("   WHERE 1=1 ");
            sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
            sb.AppendLine("  GROUP BY MACHINE ");

            if (false)//cmbEqpID.SelectedIndex < 1)
            {
                sb.AppendLine("  UNION ALL ( ");

                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'Scrap량(Kg)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN DRAIN_SCRAP_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,DRAIN_SCRAP_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 0 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , '반출량(Kg)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN R_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,R_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 1 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'Au 농도'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_AU_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 2 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'Au 회수량(g)'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_QTY END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_QTY)) AS AVG_QTY ");
                sb.AppendLine("   , 3 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");
                sb.AppendLine(" UNION ALL");
                sb.AppendLine("   SELECT ");
                sb.AppendLine("     ' ToTal' AS MACHINE ");
                sb.AppendLine("   , 'Au 회수율'  AS GUBUN ");
                for (int i = 0; i < _dtDate.Rows.Count; i++)
                {
                    sb.AppendLine(" , SUM(CASE WHEN DRAIN_DT = '" + _dtDate.Rows[i]["DT"].ToString() + "' THEN IN_AU_RATION END) AS a" + _dtDate.Rows[i]["DT"].ToString());
                }
                sb.AppendLine("   , AVG(CONVERT(float,IN_AU_RATION)) AS AVG_QTY ");
                sb.AppendLine("   , 4 AS SEQ ");
                sb.AppendLine("   FROM DRAIN ");
                sb.AppendLine("   WHERE 1=1 ");
                sb.AppendLine("     AND  DRAIN_DT BETWEEN '" + tb_fr_yyyymm.Text + "' AND '" + tb_to_yyyymm.Text + "'");

                sb.AppendLine(" ) ");
            }
            sb.AppendLine(" ) A");
            sb.AppendLine("   GROUP BY MACHINE, GUBUN , SEQ");
            sb.AppendLine("   ORDER BY MACHINE, SEQ ");

            sbDtl.AppendLine(" ORDER BY DRAIN_PLANT, DRAIN_MACHINE,  DRAIN_IN_DT");
            sb.AppendLine(";");

            sb.AppendLine(sbDtl.ToString());

            //sb.AppendLine();

            return sb.ToString();
        }



        protected void cmbScrap_SelectedIndexChanged(object sender, EventArgs e)
        {

            string sSql = "SELECT GROUP3_DESC AS ITEM_NM, GROUP3_CODE AS ITEM_CD FROM DBO.SA_SYS_CODE with(nolock) WHERE GROUP2_CODE = '" + cmbScrap.SelectedValue + "'' GROUP BY GROUP3_DESC, GROUP3_CODE";
            SetcmbVelue(sSql, cmbEqpID);

            string sSql2 = "SELECT PLANT_CD AS ITEM_CD , PLANT_DESC AS ITEM_NM FROM DBO.SA_SYS_CODE with(nolock) GROUP BY PLANT_CD, PLANT_DESC";
            SetcmbVelue(sSql2, cmbPlant);

            ClearSheet();
            ReportViewer1.Reset();
        }

        protected void cmbPlant_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbScrap.SelectedIndex > 0)
            {
                string sSql = "SELECT GROUP3_DESC AS ITEM_NM, GROUP3_CODE AS ITEM_CD FROM DBO.SA_SYS_CODE with(nolock) WHERE PLANT_CD = '" + cmbPlant.SelectedValue + "' AND GROUP2_CODE = '" + cmbScrap.SelectedValue + "' GROUP BY GROUP3_DESC, GROUP3_CODE";
                SetcmbVelue(sSql, cmbEqpID);
            }

            ClearSheet();
            ReportViewer1.Reset();
        }

        private void SetcmbVelue(string sSql, DropDownList cmb)
        {
            DataTable UNIT = fun.getData(sSql);
            if (UNIT.Rows.Count > 0)
            {
                

                DataRow dr = UNIT.NewRow();
                
                UNIT.Rows.InsertAt(dr, 0);

                cmb.DataTextField = "ITEM_NM";
                cmb.DataValueField = "ITEM_CD";
                cmb.DataSource = UNIT;
                cmb.DataBind();
            }
        }

        protected void btnExcel_Click(object sender, EventArgs e)
        {
            int index = -1;

            if (FpSpread1.Sheets[0].Rows.Count > 0)
            {
                
                for(int i=0; i< FpSpread1.Sheets.Count; i++)
                {
                    if(!FpSpread1.Sheets[i].Visible)
                    {
                        FpSpread1.Sheets[i].Visible = true;
                        index = i;
                    }
                }

                string dt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");
                //FpSpread1.SaveExcel("C:\\Scrap_"+dt+".xlsx", FarPoint.Web.Spread.Model.IncludeHeaders.BothCustomOnly);            
                //MessageBox.ShowMessage("Scrap발생정보등록[내컴퓨터 C:\\ 저장되었습니다.].", this.Page);

                System.IO.MemoryStream m_stream = new System.IO.MemoryStream();


                //Warning[] warnings;
                //string[] streamids;
                //string mimeType;
                //string encoding;
                //string extension;

                //System.IO.MemoryStream m_stream2 = new System.IO.MemoryStream();

                //byte[] bytes = ReportViewer1.LocalReport.Render(
                //   "Excel", null, out mimeType, out encoding,
                //    out extension,
                //   out streamids, out warnings);

                //m_stream2.Write(bytes, 0, bytes.Length);


                FpSpread1.SaveExcel(m_stream, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders);
                m_stream.Position = 0;

                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("content-disposition", "inline; filename=" + dt + ".xls");
                Response.BinaryWrite(m_stream.ToArray());
                //Response.BinaryWrite(m_stream2.ToArray());
                Response.End();

                if (index > -1)
                {
                    FpSpread1.Sheets[index].Visible = false;
                }
            }
        }

        protected void FpSpread1_ColumnHeaderClick(object sender, SpreadCommandEventArgs e)
        {
            DataSet ds = (DataSet)e.SheetView.DataSource;

        }

        protected void btnDtl_Click(object sender, EventArgs e)
        {
            if (FpSpread1.Sheets[0].Rows.Count > 0)
            {
                if (FpSpread1.ActiveSheetViewIndex == 0)
                {
                    FpSpread1.ActiveSheetViewIndex = 1;
                    FpSpread1.Sheets[0].Visible = false;
                    FpSpread1.Sheets[1].Visible = true;
                    btnDtl.Text = "Trend";
                }
                else
                {
                    FpSpread1.ActiveSheetViewIndex = 0;
                    FpSpread1.Sheets[0].Visible = true;
                    FpSpread1.Sheets[1].Visible = false;
                    btnDtl.Text = "상세";
                }
            }
        }

        private void ClearSheet()
        {
            for(int i =0; i< FpSpread1.Sheets.Count; i++)
            {
                FpSpread1.Sheets[i].Rows.Count =0;
                FpSpread1.Sheets[i].Columns.Count =0;
            }
        }

    }
}
