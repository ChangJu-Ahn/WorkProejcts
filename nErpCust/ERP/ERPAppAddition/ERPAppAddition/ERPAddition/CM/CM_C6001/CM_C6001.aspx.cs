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
using Microsoft.Reporting.WebForms;
using SRL.UserControls;
using System.Drawing;


namespace ERPAppAddition.ERPAddition.CM.CM_C6001
{
    public partial class CM_C6001 : System.Web.UI.Page
    {
        #region Global Variable Declaration (Sql Connection, Command, Reader)
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sqlCmd;
        SqlDataReader sqlDataReader;
        int cnt;
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                try
                {
                    InitDropDownPlant();
                    InitMultiComboAcct();
                    WebSiteCount();
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                    //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
                }
                finally
                {
                    if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                        sqlConn.Close();
                }
            }

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sqlConn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            string sqlQuery = GetSelectQuery();

            try
            {
                SqlState_Check();

                sqlCmd = new SqlCommand(sqlQuery, sqlConn);
                sqlCmd.CommandTimeout = 4000;
                da.SelectCommand = sqlCmd;
                da.SelectCommand.CommandTimeout = 4000;
                da.Fill(dt);

                dgList.DataSource = dt;

                if (dt.Rows.Count > 0) dgList.DataBind();
                else
                {
                    dgList.Controls.Clear();
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "alert('검색된 정보가 없습니다.');", true);
                }


                //ReportViewerSetting(dt);
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
            }
            finally
            {
                if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
            }
        }

        #region ReportViewer Setting
        //protected void ReportViewerSetting(DataTable dt)
        //{
        //    ReportDataSource rds = new ReportDataSource("rdsContrastBom", dt);

        //    ReportViewer1.Reset();
        //    ReportViewer1.LocalReport.ReportPath = Server.MapPath("CM_C6001.rdlc");
        //    ReportViewer1.LocalReport.DisplayName = "투입현황대비_실사용량 [" + DateTime.Now.ToShortDateString() + "]";
        //    ReportViewer1.LocalReport.DataSources.Add(rds);
        //}
        #endregion

        #region Page Controls Setting (DropDownList, MultiCheckCombo)
        protected void InitDropDownPlant()
        {
            string sqlQuery = @"Select Top 101 PLANT_CD,PLANT_NM From   B_PLANT Where  PLANT_CD>= '' order by PLANT_CD";

            txtdate_From.Text = DateTime.Now.ToString("yyyyMM");
            txtdate_To.Text = DateTime.Now.ToString("yyyyMM");

            SqlState_Check();
            sqlCmd = new SqlCommand(sqlQuery, sqlConn);
            sqlDataReader = sqlCmd.ExecuteReader();

            ddl_Plant.DataSource = sqlDataReader;
            ddl_Plant.DataValueField = "PLANT_CD";
            ddl_Plant.DataTextField = "PLANT_NM";
            ddl_Plant.DataBind();
        }

        protected void InitMultiComboAcct()
        {
            string sqlQuery = @"Select Top 101 MINOR_CD, '(' + MINOR_CD + ')_' + MINOR_NM AS MINOR_NM From   B_MINOR Where  MAJOR_CD = 'P1001' And MINOR_CD>= '' order by MINOR_CD";
            SqlState_Check();

            sqlCmd = new SqlCommand(sqlQuery, sqlConn);
            sqlDataReader = sqlCmd.ExecuteReader();
            mcc_Acct.AddItems(sqlDataReader, "MINOR_NM", "MINOR_CD");
            mcc_Acct.BackColor = Color.Yellow;
        }
        #endregion

        #region MSSQL Query (기존에는 프로시저로 동작하였으나, acct_cd를 여러개를 조회할 경우 파라메터로 배열 값을 던져야 함, 근데 프로시저 안에서 임시테이블을 사용하다 보니 exec문이 안먹을 경우가 있어 프로그램 내부로 쿼리이식)
        protected string GetSelectQuery()
        {
            StringBuilder sb = new StringBuilder();

            string sqlAcct = (mcc_Acct.SQLText.Trim() == "") ? "%" : mcc_Acct.SQLText.Trim().ToString();
            string from_Day = txtdate_From.Text.ToString();
            string To_Day = txtdate_To.Text.ToString();
            string plant_Cd = ddl_Plant.SelectedValue.ToString().ToUpper();
            string query_Item = (txtItem.Text.Length > 0) ? "AND A.ITEM_CD = '" + txtItem.Text.Trim() + "'" : "";

            //최초 with문으로 개발했으나, 임시테이블의 성능이 더 좋아 임시테이블로 변경
            sb.AppendLine(string.Format(@"
                                            SET NOCOUNT ON
                                            SET ANSI_WARNINGS OFF
                                            SET ARITHIGNORE ON
                                            SET ARITHABORT OFF

                                            IF OBJECT_ID('TEMPDB.DBO.#ERP_DATE') IS NOT NULL                                
                                            BEGIN
                                                DROP TABLE #ERP_DATE      
                                            END --  IF OBJECT_ID('TEMPDB.DBO.#ERP_DATE') IS NOT NULL
                                            
                                            SELECT Z.* INTO #ERP_DATE 
                                            FROM (
                                                    SELECT	
                                                        C.PLANT_NM
                                                        ,C.PLANT_CD
                                                        ,B.ITEM_ACCT
                                                        ,A.ITEM_CD
                                                        ,D.ITEM_NM
	                                                    ,SUM(A.INV_QTY) AS BASE_QTY
                                                        ,SUM(A.INV_AMT) AS BASE_AMT
	                                                    ,0 AS MR_QTY,0 AS MR_AMT,0 AS PR_QTY,0 AS PR_AMT,0 AS OR_QTY,0 AS OR_AMT,0 AS ST_DEB_QTY,0 AS ST_DEB_AMT
	                                                    ,0 AS PI_QTY,0 AS PI_AMT,0 AS DI_QTY,0 AS DI_AMT,0 AS OI_QTY,0 AS OI_AMT,0 AS ST_CRE_QTY,0 AS ST_CRE_AMT
                                                        FROM	I_MONTHLY_INVENTORY A(NOLOCK) 
                                                        LEFT OUTER JOIN B_ITEM_BY_PLANT B(NOLOCK)
                                                             ON A.PLANT_CD = B.PLANT_CD AND A.ITEM_CD = B.ITEM_CD
                                                        INNER JOIN B_PLANT C(NOLOCK)
                                                             ON A.PLANT_CD = C.PLANT_CD
                                                        INNER JOIN B_ITEM D(NOLOCK)
                                                             ON B.ITEM_CD = D.ITEM_CD
                                                    WHERE 1=1
                                                    AND A.PLANT_CD = '{4}'
                                                    AND	A.MNTH_INV_YEAR BETWEEN CONVERT(CHAR(4), DATEADD(DAY, -1, '{0}' + '{2}' + '01'), 112) AND CONVERT(CHAR(4), DATEADD(DAY, -1, '{1}' + '{3}' + '01'), 112)
                                                    AND	A.MNTH_INV_MONTH BETWEEN CONVERT(CHAR(2), DATEADD(DAY, -1, '{0}' + '{2}' + '01'), 110) AND CONVERT(CHAR(2), DATEADD(DAY, -1, '{1}' + '{3}' + '01'), 110)
                                                    AND	B.ITEM_ACCT IN ({5})
                                                    AND (A.INV_QTY <> 0 OR A.INV_AMT <> 0) {6}
                                                    GROUP BY C.PLANT_NM, C.PLANT_CD, B.ITEM_ACCT, A.ITEM_CD, D.ITEM_NM
                                                    UNION ALL
                                                    SELECT 
                                                          F.PLANT_NM AS PLANT_NM
                                                        , A.PLANT_CD AS PLANT_CD
                                                        , C.ITEM_ACCT AS ITEM_ACCT
                                                        , C.ITEM_CD AS ITEM_CD
                                                        , D.ITEM_NM AS ITEM_NM
                                                        , 0 AS BASE_QTY
                                                        , 0 AS BASE_AMT
                                                        , SUM(CASE WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY
                                                        WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS MR_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS MR_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY
                                                        WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS PR_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS PR_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY
                                                        WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS OR_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS OR_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY ELSE 0 END) AS ST_DEB_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT ELSE 0 END ) AS ST_DEB_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY 
                                                        WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS PI_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS PI_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY 
                                                        WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS DI_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS DI_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY 
                                                        WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS OI_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT
                                                        WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS OI_AMT,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY ELSE 0 END) AS ST_CRE_QTY,
                                                        SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT ELSE 0 END ) AS ST_CRE_AMT
                                                    FROM I_GOODS_MOVEMENT_DETAIL A (nolock), I_GOODS_MOVEMENT_HEADER B (nolock), B_ITEM_BY_PLANT C (nolock), 
                                                    B_ITEM D (nolock), B_PLANT F (nolock) 
                                                    WHERE C.PLANT_CD = A.PLANT_CD
                                                    AND A.ITEM_CD = C.ITEM_CD
                                                    AND B.DOCUMENT_YEAR = A.DOCUMENT_YEAR
                                                    AND B.ITEM_DOCUMENT_NO = A.ITEM_DOCUMENT_NO
                                                    AND A.PLANT_CD  = F.PLANT_CD 
                                                    AND C.ITEM_CD = D.ITEM_CD
                                                    AND A.DELETE_FLAG = 'N'
                                                    AND A.DOCUMENT_YEAR BETWEEN '{0}' AND '{1}'
                                                    AND convert(char(6), B.DOCUMENT_DT, 112) BETWEEN '{0}'+'{2}' AND '{1}'+'{3}'
                                                    AND C.PLANT_CD = '{4}'
                                                    AND C.ITEM_ACCT IN ({5}) {6} 
                                                    GROUP BY F.PLANT_NM,A.PLANT_CD,C.ITEM_ACCT, C.ITEM_CD, D.ITEM_NM
                                                ) Z

                                         SELECT 
                                             PLANT_CD
                                        	,PLANT_NM
                                        	,ITEM_ACCT
                                        	,ITEM_CD
                                        	,ITEM_NM
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(BASE_QTY), 0)) AS BASE_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(BASE_AMT), 0)) AS BASE_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(MR_QTY), 0)) AS MR_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(MR_AMT), 0)) AS MR_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(PR_QTY), 0)) AS PR_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(PR_AMT), 0)) AS PR_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(OR_QTY), 0)) AS OR_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(OR_AMT), 0)) AS OR_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(ST_DEB_QTY), 0)) AS ST_DEB_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(ST_DEB_AMT), 0)) AS ST_DEB_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(PI_QTY), 0)) AS PI_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(PI_AMT), 0)) AS PI_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(DI_QTY), 0)) AS DI_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(DI_AMT), 0)) AS DI_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(OI_QTY), 0)) AS OI_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(OI_AMT), 0)) AS OI_AMT
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(ST_CRE_QTY), 0)) AS ST_CRE_QTY
                                        	,CONVERT(DECIMAL(16,2), ISNULL(SUM(ST_CRE_AMT), 0)) ST_CRE_AMT
                                        	,CONVERT(DECIMAL(16,2), (ISNULL(SUM(BASE_QTY), 0) + ISNULL(SUM(MR_QTY), 0) + ISNULL(SUM(PR_QTY), 0) + ISNULL(SUM(OR_QTY), 0) + ISNULL(SUM(ST_DEB_QTY), 0)) - (ISNULL(SUM(PI_QTY), 0) + ISNULL(SUM(DI_QTY), 0) + ISNULL(SUM(OI_QTY), 0) + ISNULL(SUM(ST_CRE_QTY), 0))) AS NEXT_QTY
                                        	,CONVERT(DECIMAL(16,2), (ISNULL(SUM(BASE_AMT), 0) + ISNULL(SUM(MR_AMT), 0) + ISNULL(SUM(PR_AMT), 0) + ISNULL(SUM(OR_AMT), 0) + ISNULL(SUM(ST_DEB_AMT), 0)) - (ISNULL(SUM(PI_AMT), 0) + ISNULL(SUM(DI_AMT), 0) + ISNULL(SUM(OI_AMT), 0) + ISNULL(SUM(ST_CRE_AMT), 0))) AS NEXT_AMT
                                            FROM #ERP_DATE 
                                            GROUP BY PLANT_CD
	                                            ,PLANT_NM
	                                            ,ITEM_ACCT
	                                            ,ITEM_CD
	                                            ,ITEM_NM  ", from_Day.Substring(0, 4), To_Day.Substring(0, 4), from_Day.Substring(4, 2), To_Day.Substring(4, 2), plant_Cd, sqlAcct, query_Item)
                );

            return sb.ToString();

        }
        #endregion

        protected void SqlState_Check()
        {
            if (sqlConn == null || sqlConn.State == ConnectionState.Closed)
                sqlConn.Open();

            if (sqlDataReader != null && sqlDataReader.FieldCount > 0)
                sqlDataReader.Close();
        }

        protected void dgList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataRowView rowView = (DataRowView)e.Row.DataItem;
            string[] tempArray;

            if (rowView == null)
                cnt = 0;
            else
            {
                tempArray = new string[rowView.Row.ItemArray.Length];

                for (int i = 0; i < tempArray.Length; i++)
                {
                    if ((cnt % 2) == 0) e.Row.BackColor = Color.FromName("#D5D5D5");
                    e.Row.Cells[i].Text = rowView[i].ToString();
                }

                cnt += 1;
            }
        }

        protected void btnExcelDown_Click(object sender, EventArgs e)
        {
            Response.Clear();
            //파일이름 설정
            string fName = string.Format("{0}.xls", DateTime.Now.ToString("yyyyMMddHHmmss"));
            //헤더부분에 내용을 추가
            Response.AddHeader("Content-Disposition", "attachment;filename=" + fName);
            Response.Charset = "utf-8";
            //컨텐츠 타입 설정
            string encoding = Request.ContentEncoding.HeaderName;
            Response.ContentType = "application/ms-excel";
            Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=" + encoding + "'>");

            System.IO.StringWriter SW = new System.IO.StringWriter();
            HtmlTextWriter HW = new HtmlTextWriter(SW);
            SW.WriteLine(" "); //한글 깨짐 방지

            dgList.RenderControl(HW);
            Response.Write(SW.ToString());
            Response.End();
            HW.Close();
            SW.Close();
        }

        public override void VerifyRenderingInServerForm(System.Web.UI.Control control)
        {
            // Confirms that an HtmlForm control is rendered for the specified ASP.NET server control at run time.
        }


    }
}