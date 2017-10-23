using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using Microsoft.Reporting.WebForms;

namespace ERPAppAddition.ERPAddition.MM.MM_MM002
{
    public partial class MM_MM002 : System.Web.UI.Page
    {
        SqlConnection _sqlConn;
        SqlCommand _sqlCmd;
        SqlDataReader _sqlDataReader;
        public string connDBnm = string.Empty;

        protected void Page_Init(object sender, EventArgs e)
        {
            InitParam();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    InitDropDownPlant();
                    WebSiteCount();

                    TimeSpan ts = new TimeSpan(-7, 0, 0, 0);
                    
                    cal_From.SelectedDate = DateTime.Now.Date.Add(ts);
                    cal_to.SelectedDate = DateTime.Now.Date;
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}')", ex.Message.ToString()), true);
                }
                finally
                {
                    if (_sqlConn != null || _sqlConn.State == ConnectionState.Open)
                        _sqlConn.Close();
                }
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = _sqlConn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void InitDropDownPlant()
        {
            string sqlQuery = @"Select Top 101 BIZ_AREA_NM, BIZ_AREA_CD From B_BIZ_AREA";

            SqlStateCheck();
            _sqlCmd = new SqlCommand(sqlQuery, _sqlConn);
            _sqlDataReader = _sqlCmd.ExecuteReader();

            ddl_BIZ_AREA.DataSource = _sqlDataReader;
            ddl_BIZ_AREA.DataValueField = "BIZ_AREA_CD";
            ddl_BIZ_AREA.DataTextField = "BIZ_AREA_NM";
            ddl_BIZ_AREA.DataBind();
        }

        #region DB Connection and DataReader State Check
        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
                _sqlDataReader.Close();
        }
        #endregion

        protected void InitParam()
        {
            if (Request.QueryString["dbName"] != null && Request.QueryString["dbName"].ToString() != "")
                connDBnm = Request.QueryString["dbName"].ToString();
            else
                connDBnm = "nepes";

            hdndbnm.Value = connDBnm;
            _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings[connDBnm].ConnectionString);

            if (Request.QueryString["userId"] != null && Request.QueryString["userId"].ToString() != "")
                ViewState["userId"] = Request.QueryString["userId"].ToString();
            else
                ViewState["userId"] = "DEV";

            lblerpName.Text = _sqlConn.Database.ToString().ToUpper();
        }

        protected string GetSqlQuery()
        {
            StringBuilder sbBaseData = new StringBuilder();
            StringBuilder sbMainQuery = new StringBuilder();

            sbBaseData.AppendLine(@"
                	        SELECT 
	                           A.BIZ_AREA
	                          ,G.BIZ_AREA_NM
	                          ,F.ID_DT
	                          ,A.BAS_NO
	                          ,F.DISCHGE_PORT+ ' (' + H.MINOR_NM + ')' AS 'DISCHGE_PORT'
	                          ,F.TRANSPORT+ ' (' + I.MINOR_NM + ')' AS 'TRANSPORT'
	                          ,F.IP_NO
	                          ,C.JNL_CD 
	                          ,C.JNL_NM 
	                          ,A.CURRENCY
	                          ,A.CHARGE_DOC_AMT
	                          ,A.VAT_DOC_AMT
	                          ,A.CHARGE_LOC_AMT
	                          ,A.VAT_LOC_AMT
	                          ,A.PAYEE_CD
	                          ,E.BP_NM
	                        FROM M_PURCHASE_CHARGE A RIGHT JOIN B_MINOR B ON B.MINOR_CD = A.PROCESS_STEP AND B.MAJOR_CD = 'M9014'
		                         RIGHT JOIN A_JNL_ITEM C ON C.JNL_CD = A.CHARGE_TYPE AND C.JNL_TYPE = 'EC'
		                         RIGHT JOIN B_PUR_GRP D ON D.PUR_GRP = A.PUR_GRP
		                         RIGHT JOIN B_BIZ_PARTNER E ON E.BP_CD = A.PAYEE_CD
		                         RIGHT JOIN M_CC_HDR F ON A.BAS_NO = F.BL_NO
		                         RIGHT JOIN B_BIZ_AREA G ON A.BIZ_AREA = G.BIZ_AREA_CD 
		                         RIGHT JOIN B_MINOR H ON F.DISCHGE_PORT = H.MINOR_CD AND H.MAJOR_CD = 'B9092' 
		                         RIGHT JOIN B_MINOR I ON F.TRANSPORT = I.MINOR_CD AND I.MAJOR_CD = 'B9009'  
	                        WHERE 1=1
                         ");

            if (ddl_BIZ_AREA.SelectedValue.Length > 0) //사업부 검색조건
                sbBaseData.AppendLine(string.Format("AND A.BIZ_AREA = '{0}'", ddl_BIZ_AREA.SelectedValue.ToString()));

            if (hdnPartner.Value.Length > 0) //업체 검색조건
                sbBaseData.AppendLine(string.Format("AND A.BP_CD = '{0}'", hdnPartner.Value.ToString()));

            if (txtBL.Text.Length > 0) //BL 검색조건
                sbBaseData.AppendLine(string.Format("AND A.BAS_NO = '{0}'", txtBL.Text.ToString()));

            if (txtdate_From.Text.Length > 0) //통관일자 검색조건(From)
                sbBaseData.AppendLine(string.Format("AND F.ID_DT >= '{0}'", txtdate_From.Text.ToString()));

            if (txtdate_To.Text.Length > 0) //통관일자 검색조건(To) 
                sbBaseData.AppendLine(string.Format("AND F.ID_DT <= '{0}'", txtdate_To.Text.ToString()));

            sbMainQuery.AppendLine(
                    string.Format(@"
                    WITH BaseData AS 
                    (
                        {0}
                    )
                    , DataGrouping AS 
                    (
	                    SELECT 
		                    A.BIZ_AREA_NM
	                       ,A.ID_DT
	                       ,A.BAS_NO
	                       ,A.DISCHGE_PORT
	                       ,A.TRANSPORT
	                       ,A.IP_NO
	                       ,A.JNL_CD
	                       ,CASE WHEN A.JNL_CD = 'SPC05' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC05_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC05' THEN SUM(A.VAT_LOC_AMT) ELSE 0 END AS 'SPC05_VAT_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC03' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC03_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC03' THEN SUM(A.VAT_LOC_AMT) ELSE 0 END AS 'SPC03_VAT_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC14' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC14_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC14' THEN SUM(A.VAT_LOC_AMT) ELSE 0 END AS 'SPC14_VAT_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC02' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC02_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC02' THEN SUM(A.VAT_LOC_AMT) ELSE 0 END AS 'SPC02_VAT_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC12' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC12_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC12' THEN SUM(A.VAT_LOC_AMT) ELSE 0 END AS 'SPC12_VAT_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC01' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC01_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC19' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC19_SUM'
	                       ,CASE WHEN A.JNL_CD = 'SPC18' THEN SUM(A.CHARGE_LOC_AMT) ELSE 0 END AS 'SPC18_SUM'
	                    FROM BaseData A
	                    GROUP BY 
		                    A.BIZ_AREA_NM
	                       ,A.ID_DT
	                       ,A.BAS_NO
	                       ,A.DISCHGE_PORT
	                       ,A.TRANSPORT
	                       ,A.IP_NO
	                       ,A.JNL_CD
                    )

                    --관리항목별 산출된 금액을 한 개의 Row로 병합
                SELECT 
                  *
                  FROM
                  (
                    SELECT 
	                    Z.BIZ_AREA_NM
                       ,CONVERT(VARCHAR, Z.ID_DT, 23) AS 'ID_DT'
                       ,Z.BAS_NO
                       ,Z.DISCHGE_PORT
                       ,Z.TRANSPORT
                       ,Z.IP_NO
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC05_SUM)), 1) AS 'SPC05'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC03_SUM)), 1) AS 'SPC03'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC14_SUM)), 1) AS 'SPC14'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC02_SUM)), 1) AS 'SPC02'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC12_SUM)), 1)  AS 'SPC12'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC01_SUM)), 1) AS 'SPC01'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC19_SUM)), 1) AS 'SPC19'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC18_SUM)), 1) AS 'SPC18'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC05_VAT_SUM)), 1) AS 'SPC05_VAT'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC03_VAT_SUM)), 1) AS 'SPC03_VAT'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC14_VAT_SUM)), 1) AS 'SPC14_VAT'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC02_VAT_SUM)), 1) AS 'SPC02_VAT'
                       ,CONVERT(VARCHAR, CONVERT(MONEY, SUM(SPC12_VAT_SUM)), 1) AS 'SPC12_VAT'
                       ,'0' AS SRT
                    FROM DataGrouping Z
                    GROUP BY 
	                    Z.BIZ_AREA_NM
                       ,Z.ID_DT
                       ,Z.BAS_NO
                       ,Z.DISCHGE_PORT
                       ,Z.TRANSPORT
                       ,Z.IP_NO", sbBaseData.ToString())
             );

            sbMainQuery.AppendLine(string.Format(@"UNION ALL
                SELECT 
                          BIZ_AREA_NM 
                          ,CONVERT(VARCHAR,TEMP_GL_DT, 23) AS 'ID_DT'
                          , TEMP_GL_DESC + GRP 
                          --, TEMP_GL_NO
                          , '' AS DISCHGE_PORT
                          , '' AS TRANSPORT
                          , TEMP_GL_NO AS IP_NO
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43001113', '53011713', '5301471713') THEN SUM(ITEM_AMT) END,0)) AS 'SPC05'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002302', '53012902', '53012910') THEN SUM(ITEM_AMT) END,0)) AS 'SPC03'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '') THEN SUM(ITEM_AMT) END,0)) AS 'SPC14'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002306', '53012906', '5301472906') THEN SUM(ITEM_AMT) END,0)) AS 'SPC02'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002305', '53012905', '5301472905') THEN SUM(ITEM_AMT) END,0)) AS 'SPC12'
                          ,'' AS SPC01
                          ,'' AS SPC19
                          ,'' AS SPC18
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43001113', '53011713', '5301471713') THEN SUM(VAT_AMT) END,0)) AS 'SPC05_VAT'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002302', '53012902', '53012910') THEN SUM(VAT_AMT) END,0)) AS 'SPC03_VAT'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '') THEN SUM(VAT_AMT) END,0)) AS 'SPC14_VAT'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002306', '53012906', '5301472906') THEN SUM(VAT_AMT) END,0)) AS 'SPC02_VAT'
                          , CONVERT(VARCHAR,ISNULL(CASE WHEN ACCT_CD IN ( '43002305', '53012905', '5301472905') THEN SUM(VAT_AMT) END,0)) AS 'SPC12_VAT'
                          ,'1' AS SRT
                        FROM
                          (
  
                          SELECT 
	                        GL.BIZ_AREA_CD
                           ,BA.BIZ_AREA_NM
                           ,GL.TEMP_GL_DT 
                           ,GL.ISSUED_DT 
                           ,TEMP_GL_DESC 
                           ,GL.TEMP_GL_NO
                           ,GL_ITM.ACCT_CD 
                           ,GL_ITM.ITEM_AMT 
                           ,GL_ITM.ITEM_LOC_AMT 
                           ,GL_ITM.VAT_AMT  
                           ,GL_ITM.VAT_LOC_AMT  
                           , CASE WHEN GL_ITM.ACCT_CD IN ('43001113', '43002302', '43002305', '43002306') THEN '[제]'
                                 WHEN GL_ITM.ACCT_CD IN ('53011713', '53012902', '53012905', '53012906', '53012910') THEN '[판]'
                                 WHEN GL_ITM.ACCT_CD IN ('5301471713', '5301472905', '5301472906') THEN '[경상]'
                            END AS GRP
                           FROM A_TEMP_GL GL WITH(NOLOCK)
                          RIGHT JOIN  A_TEMP_GL_ITEM GL_ITM WITH(NOLOCK)
                          ON GL.TEMP_GL_NO = GL_ITM.TEMP_GL_NO
                          RIGHT JOIN B_BIZ_AREA BA WITH(NOLOCK)
                          ON GL.BIZ_AREA_CD = BA.BIZ_AREA_CD
                          WHERE GL_ITM.DR_CR_FG = 'DR'
                          AND ACCT_CD IN ( 
                                          '43001113' --제)세금과공과(관세)
										 ,'43002302'  --제)운반보관료(수입운임)
										 ,'43002305'  --제)운반보관료(통관수수료)
										 ,'43002306'  --제)운반보관료(육상운임)
										 ,'53011713'  --판)세금과공과(관세)
										 ,'53012902'  --판)운반보관료(수입운임)
										 ,'53012905'  --판)운반보관료(통관)
										 ,'53012906'  --판)운반보관료(육상운임)
										 ,'53012910'  --판)운반보관료(수출운임)
										 ,'5301471713'  --경상)세금과공과(관세)
										 ,'5301472905'  --경상)운반보관료(통관수수료)
										 ,'5301472906'--경상)운반보관료(육상운임)
                        )
                        ) A
                        WHERE 1=1 "));
            if (ddl_BIZ_AREA.SelectedValue.Length > 0) //사업부 검색조건
                sbMainQuery.AppendLine(string.Format("AND A.BIZ_AREA_CD = '{0}'", ddl_BIZ_AREA.SelectedValue.ToString()));

            if (hdnPartner.Value.Length > 0) //업체 검색조건
                sbMainQuery.AppendLine("AND 1 = 2");

            if (txtBL.Text.Length > 0) //BL 검색조건
                sbMainQuery.AppendLine("AND 1 = 2");

            if (txtdate_From.Text.Length > 0) //통관일자 검색조건(From)
                sbMainQuery.AppendLine(string.Format("AND TEMP_GL_DT >= '{0}'", txtdate_From.Text.ToString()));

            if (txtdate_To.Text.Length > 0) //통관일자 검색조건(To) 
                sbMainQuery.AppendLine(string.Format("AND TEMP_GL_DT <= '{0}'", txtdate_To.Text.ToString()));


            sbMainQuery.AppendLine(string.Format(@"
			GROUP BY BIZ_AREA_NM 
                          , TEMP_GL_DT 
                          , TEMP_GL_DESC + GRP
                          , ACCT_CD
                          , TEMP_GL_NO)A
            ORDER BY SRT, ID_DT, BAS_NO"));

            return sbMainQuery.ToString();
        }


        protected void btnSelect_Click(object sender, EventArgs e)
        {
            DataTable dtReturn = new DataTable();
            SqlDataAdapter Sqladapter;

            string sqlQuery = GetSqlQuery();

            try
            {
                SqlStateCheck();
                Sqladapter = new SqlDataAdapter(sqlQuery, _sqlConn);
                Sqladapter.Fill(dtReturn);

                SetReportView(dtReturn);
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}')", ex.Message.ToString()), true);
            }
            finally
            {
                if (_sqlConn != null || _sqlConn.State == ConnectionState.Open)
                    _sqlConn.Close();
            }
        }

        protected void SetReportView(DataTable dt)
        {
            ReportViewer1.Reset();

            ReportViewer1.LocalReport.ReportPath = Server.MapPath("MM_MM002_RP.rdlc");
            ReportViewer1.LocalReport.DisplayName = "물류비용조회" + ddl_BIZ_AREA.Text.ToString() + DateTime.Now.ToShortDateString();

            ReportDataSource rds = new ReportDataSource();
            rds.Name = "DataSet1";
            rds.Value = dt;
            ReportViewer1.LocalReport.DataSources.Add(rds);

            ReportViewer1.LocalReport.Refresh();
        }

    }
}