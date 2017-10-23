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


namespace ERPAppAddition.ERPAddition.CM.CM_C4001
{
    public partial class CM_C4001 : System.Web.UI.Page
    {
        #region Global Variable Declaration (Sql Connection, Command, Reader)
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sqlCmd;
        SqlDataReader sqlDataReader;
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
                    MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
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

                sqlCmd = new SqlCommand(sqlQuery,sqlConn);
                da.SelectCommand = sqlCmd;
                da.Fill(dt);

                ReportViewerSetting(dt);
            }
            catch (Exception ex)
            {
                MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
            }
            finally
            {
                if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
            }
        }

        #region ReportViewer Setting
        protected void ReportViewerSetting(DataTable dt)
        {
            ReportDataSource rds = new ReportDataSource("rdsContrastBom", dt);

            ReportViewer1.Reset();
            ReportViewer1.LocalReport.ReportPath = Server.MapPath("CM_C4001.rdlc");
            ReportViewer1.LocalReport.DisplayName = "투입현황대비_실사용량 [" + DateTime.Now.ToShortDateString() + "]";
            ReportViewer1.LocalReport.DataSources.Add(rds);
        }
        #endregion

        #region Page Controls Setting (DropDownList, MultiCheckCombo)
        protected void InitDropDownPlant()
        {
            string sqlQuery = @"Select Top 101 PLANT_CD,PLANT_NM From   B_PLANT Where  PLANT_CD>= '' order by PLANT_CD";
           
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

            sb.AppendLine(@"
                    SET ANSI_WARNINGS OFF 
                    SET ARITHIGNORE ON                        
                    SET ARITHABORT OFF    

                    BEGIN TRY                        
                     DROP TABLE #TEMP_TBL_ITEM                        
                     DROP TABLE #TEMP_TBL_RECEIPT                        
                    END TRY                        
                    BEGIN CATCH                        
                     PRINT '임시테이블 없음'                        
                    END CATCH  

                     SELECT B.* INTO #TEMP_TBL_ITEM FROM (                        
                        select case when grouping(a.child_plant_cd)=1 then '%1' else a.child_plant_cd end plant_cd,  -- 공장                           
                            case when grouping(a.cost_cd)=1  then '%2' else a.cost_cd end cost_cd,   -- C/C                          
                            case when grouping(a.cost_cd)=1  then '' else max(b.cost_nm) end cost_nm,  -- C/C명                          
                            case when grouping(a.order_no)=1 then '%3' else a.order_no end order_no,   -- 오더번호                          
                            case when grouping(a.po_seq_no)=1 then '' else a.po_seq_no end  po_seq_no ,  -- 오더SEQ                          
                            case when grouping(a.wc_cd)=1  then '%4' else a.wc_cd end wc_cd,   -- 작업장                          
                            case when grouping(a.wc_cd)=1  then '' else max( isnull(c.wc_nm,'') + isnull(d.pur_grp_nm,'') )end  wc_nm, -- 작업장명                     
                            case when grouping(a.child_item_acct)=1 then '%5' else a.child_item_acct end item_acct,  -- Hidden                          
                            case when grouping(a.child_item_acct)=1 then '' else max(e.minor_nm) end acct_nm,  -- 품목계정                          
                            case when grouping(a.child_item_cd)=1 then '%6' else a.child_item_cd end item_cd,  -- 자품목                          
                            case when grouping(a.child_item_cd)=1 then '' else max( f.item_nm ) end item_nm,  -- 자품목명                          
                            case when grouping(a.mov_type)=1 then '' else a.mov_type end  mov_type,   -- 수불유형                          
                            case when grouping(a.mov_type)=1 then '' else  max(isnull(g.minor_nm,'')) end mov_nm, -- 수불유형명                          
                            sum(a.this_wip_qty) as this_wip_qty, -- 투입수량                          
                            sum(a.this_wip_amt) as this_wip_amt -- 투입금액                          
                         from c_bom_issue_by_opr_s a(nolock)                          
                             join b_cost_center  b(nolock) on a.cost_cd = b.cost_cd                          
                             left outer join p_work_center c(nolock) on a.wc_cd = c.wc_cd                          
                             left outer join b_pur_grp d(nolock) on a.wc_cd = d.pur_grp                          
                             join b_minor  e(nolock) on a.child_item_acct = e.minor_cd and e.major_cd = 'P1001'                          
                             join b_item  f(nolock) on a.child_item_cd = f.item_cd                          
                             join b_minor  g(nolock) on a.mov_type = g.minor_cd and g.major_cd = 'I0001'   
                ");

            sb.AppendLine("where a.yyyymm between '" + from_Day + "' and '" + To_Day + "'");
            sb.AppendLine("and a.child_plant_cd = '" + plant_Cd + "'");

            if (sqlAcct != "%")
                sb.AppendLine("and a.child_item_acct IN (" + sqlAcct + ")");

            sb.AppendLine(@"group by a.child_plant_cd,a.cost_cd,a.order_no,a.po_seq_no,a.wc_cd,a.child_item_acct,a.child_item_cd,a.mov_type                        
                                having (sum(a.this_wip_qty) <> 0 or sum(a.this_wip_amt) <> 0)                          
                            and (grouping(a.child_plant_cd)+grouping(a.cost_cd)+grouping(a.order_no)+grouping(a.po_seq_no) <> 1) 
                            and (grouping(a.child_plant_cd)+grouping(a.cost_cd)+grouping(a.order_no)+grouping(a.po_seq_no)+grouping(a.wc_cd)+grouping(a.child_item_acct)+grouping(a.child_item_cd)+grouping(a.mov_type) <> 1) -- 수불유형 합계 제외                          
                            ) B");


            sb.AppendLine(@"SELECT A.* INTO #TEMP_TBL_RECEIPT FROM (  
                                SELECT temp.mov_type, temp.ITEM_CD,temp.ITEM_ACCT,temp.PLANT_CD,temp.MR_QTY, temp.PR_QTY, temp.OR_QTY, temp.ST_DEB_QTY, temp.PI_QTY, temp.DI_QTY  
                                    , temp.OI_QTY, temp.ST_CRE_QTY, temp.MR_AMT, temp.minor_nm, temp.PR_AMT, temp.OR_AMT, temp.ST_DEB_AMT, temp.PI_AMT, temp.DI_AMT, temp.OI_AMT, temp.ST_CRE_AMT  
                                    FROM (SELECT A.MOV_TYPE, D.MINOR_NM, F.PLANT_NM,A.PLANT_CD,C.ITEM_ACCT, C.ITEM_CD,  
                                    SUM(CASE WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY  
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
                                FROM I_GOODS_MOVEMENT_DETAIL A (NOLOCK), I_GOODS_MOVEMENT_HEADER B (NOLOCK), B_ITEM_BY_PLANT C (NOLOCK), B_MINOR D (NOLOCK), B_MAJOR E (NOLOCK), B_PLANT F (NOLOCK)  
                                WHERE C.PLANT_CD = A.PLANT_CD  
                                    AND A.ITEM_CD = C.ITEM_CD AND A.DELETE_FLAG = 'N' AND B.DOCUMENT_YEAR = A.DOCUMENT_YEAR  
                                    AND B.ITEM_DOCUMENT_NO = A.ITEM_DOCUMENT_NO AND D.MAJOR_CD = E.MAJOR_CD AND A.MOV_TYPE = D.MINOR_CD AND A.PLANT_CD  = F.PLANT_CD
                            ");

            sb.AppendLine("AND D.MAJOR_CD = 'I0001' AND A.DOCUMENT_YEAR BETWEEN LEFT('" + from_Day + "', 4) AND LEFT('" + To_Day + "', 4)");
            sb.AppendLine("AND convert(char(6), B.DOCUMENT_DT, 112) BETWEEN '" + from_Day + "' AND '" + To_Day + "'");
            sb.AppendLine("GROUP BY A.MOV_TYPE, D.MINOR_NM, F.PLANT_NM,A.PLANT_CD,C.ITEM_ACCT, C.ITEM_CD) temp");
            sb.AppendLine("        WHERE ( ( temp.PLANT_CD =  '" + plant_Cd + "'");

            if (sqlAcct != "%")
                sb.AppendLine(" AND temp.ITEM_ACCT IN (" + sqlAcct + ")");

            sb.AppendLine(@"  )) ) A  
                                  SELECT 
                                       ROW_NUMBER() OVER (ORDER BY TEMP.PLANT_CD) AS 'No'   --count                       
                                     , TEMP.PLANT_CD      AS 'PLANT_CD'                     --공장
                                     , TEMP.ITEM_ACCT  AS 'ACCT_CD'                         --품목계정
                                     , TEMP.ITEM_CD    AS 'CHILD_ITEM_CD'                   --자품목코드       
                                     , TEMP.ITEM_NM    AS 'CHILD_ITEM_NM'                   --자품목명                
                                     , CONVERT(NUMERIC(14, 2), SUM(TEMP.this_wip_qty)) AS 'INSERT_QTY'   -- 투입수량
                                     , CONVERT(NUMERIC(14, 0), SUM(TEMP.this_wip_amt)) AS 'INSERT_AMT'   -- 투입금액                       
                                     , CONVERT(NUMERIC(14, 2), SUM(TEMP.usage_qty)) AS 'USAGE_QTY'   -- 사용량           
                                     , CONVERT(NUMERIC(14, 0), SUM(TEMP.usage_amt)) AS 'USAGE_AMT'   -- 사용금액                     
                                     , CONVERT(NVARCHAR, ISNULL(CONVERT(NUMERIC(14, 2), SUM(TEMP.usage_qty) / SUM(TEMP.this_wip_qty)), 0)) + '%' as 'QTY_CONTRAST_BOM'  --BOM대비수량(사용량/투입수량)            
                                     , CONVERT(NUMERIC(14, 0), SUM(TEMP.usage_amt) - SUM(TEMP.this_wip_amt)) as 'AMT_CONTRAST_BOM'                                      --BOM대비금액(사용금액-투입금액)                    
                                   FROM                           
                                   (                          
                                    select                           
                                     a.plant_cd            --공장                          
                                   , a.item_acct        --품목계정                          
                                   , a.item_cd        --자품목 코드                          
                                   , a.item_nm        --자품목 이름                          
                                   , sum(a.this_wip_qty) as this_wip_qty  --투입수량                          
                                   , sum(a.this_wip_amt) as this_wip_amt  --투입수량                          
                                   , 0 as usage_qty                          
                                   , 0 as usage_amt                          
                                   , 0 AS bom_qty                          
                                   , 0 AS bom_amt                          
                                    from #TEMP_TBL_ITEM a                          
                                    group by a.plant_cd , a.item_acct, a.item_cd, a.item_nm                          
                                    union all                          
                                    select                           
                                     b.plant_cd                           
                                   , d.ITEM_ACCT                          
                                   , b.item_cd                          
                                   , c.ITEM_NM                          
                                   , 0                          
                                   , 0                          
                                   , SUM(b.pi_qty + b.oi_qty) as usage_qty                          
                                   , SUM(b.pi_amt + b.oi_amt) as usage_amt                          
                                   , 0 AS bom_qty                          
                                   , 0 AS bom_amt                          
                                    from #TEMP_TBL_RECEIPT b                           
                                   inner join B_ITEM c on b.item_cd = c.ITEM_CD                           
                                   inner join B_ITEM_BY_PLANT d on b.item_cd = d.ITEM_CD and b.plant_cd = d.PLANT_CD                          
                                    group by b.plant_cd,b.item_cd, d.ITEM_ACCT, b.item_cd, c.ITEM_NM                          
                                   ) TEMP                          
                                   group by TEMP.PLANT_CD, TEMP.ITEM_ACCT, TEMP.ITEM_CD, TEMP.ITEM_NM                 
                            ");

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

    }
}