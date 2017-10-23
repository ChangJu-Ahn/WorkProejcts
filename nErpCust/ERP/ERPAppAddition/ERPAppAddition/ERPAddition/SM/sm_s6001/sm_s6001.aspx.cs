using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
//using System.Data.OleDb;
//using System.Data.OracleClient;
using System.IO;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;



namespace ERPAppAddition.ERPAddition.SM.sm_s6001
{
    public partial class sm_s6001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];                

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleConnection conn_if = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_UNIERP"].ConnectionString);

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();        

        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr, ndr, dr3;        
        DataSet ds = new DataSet();        
        DataTable dtYYYYMM= new DataTable();

        string userid = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                /*달력셋*/
                setMonth();

                /*사용자 id불러오기 최초 tab page에서 전달*/
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;
                WebSiteCount();
                //MessageBox.ShowMessage(Session["User"].ToString(), this.Page);                
            }         
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void setMonth(){
            conn.Open();            
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select PLAN_YEAR || lpad(PLAN_MONTH, 2, 0) AS M_MONTH   \n");
                sbSQL.Append("  from CALENDAR                                         \n");
                sbSQL.Append(" where PLANT = 'CCUBEDIGITAL'                           \n");
                sbSQL.Append("   and PLAN_YEAR >= '2015'                              \n");
                sbSQL.Append("   and PLAN_YEAR <= to_char(sysdate, 'yyyy')            \n");
                sbSQL.Append("   --and LPAD(PLAN_MONTH, 2, 0) <= to_char(sysdate, 'mm') \n");
                sbSQL.Append(" group by PLAN_YEAR, PLAN_MONTH                         \n");
                sbSQL.Append(" order by 1 DESC                                             \n");
                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                dr = cmd2.ExecuteReader();

                if (dr.RowSize > 0)
                {
                    tb_yyyymm.DataSource = dr;
                    tb_yyyymm.DataValueField = "M_MONTH";
                    tb_yyyymm.DataTextField = "M_MONTH";
                    tb_yyyymm.DataBind();
                    tb_yyyymm.SelectedValue = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00");  
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            //return dr;
        }      

        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            // 프로시져 실행: 기본데이타 생성
            conn.Open();                                          
            
            if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null){
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"조회년도를 입력해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }
                
            else
            {
                try
                {
                    /*막대 꺽은선 그래프*/
                    OracleCommand cmd1 = new OracleCommand(getSQL(), conn);
                    dr = cmd1.ExecuteReader();
                    ndr = cmd1.ExecuteReader();
                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s6001.rdlc");
                    ReportViewer1.LocalReport.DisplayName = tb_yyyymm.Text + "_진도현황_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dr;
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    /*원 그래프*/
                    string[] grp = { "DDI", "WLP", "FOWLP" };
                    for(int i = 0; i < grp.Length; i++ )
                    {
                        OracleCommand cmd2 = new OracleCommand(getSQL2(grp[i].ToString()), conn);
                        OracleDataReader dr2 = cmd2.ExecuteReader();
                        ReportDataSource rds2 = new ReportDataSource();
                        rds2.Name = "DataSet2_" + (i+1).ToString();                        
                        rds2.Value = dr2;
                        ReportViewer1.LocalReport.DataSources.Add(rds2);
                    }

                    

                    /*달성율, 진척율, 일경과율*/
                    OracleCommand cmd3 = new OracleCommand(getSQL3(), conn);
                    dr3 = cmd3.ExecuteReader();
                    ReportDataSource rds3 = new ReportDataSource();
                    rds3.Name = "DataSet3";
                    rds3.Value = dr3;
                    ReportViewer1.LocalReport.DataSources.Add(rds3);
                    ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x


                    DataTable dt = new DataTable();
                    dt.Load(ndr);
                    if (dt.Rows.Count > 0)
                    {
                        ReportViewer1.LocalReport.Refresh();

                        /*report view 파일떨구기*/
                        //Warning[] warnings;
                        //string[] streamids;
                        //string mimeType = string.Empty;
                        //string encoding = string.Empty;
                        //string extension = string.Empty;
                        //byte[] bytes = ReportViewer1.LocalReport.Render("pdf", null, out mimeType, out encoding, out extension, out streamids, out warnings);
                        //FileStream fs = new FileStream(@"c:\output.pdf", FileMode.Create, FileAccess.Write);
                        //fs.Write(bytes, 0, bytes.Length);
                        //fs.Close();
                    }
                    else
                    {
                        string script = "alert(\"조회된 데이터가 없습니다.\");";
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                        ReportViewer1.Reset();
                        return;
                    }

                }
                catch (Exception ex)
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
            }
        }
        

        private string getSQL()
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	SELECT                  										\n");
            sbSQL.Append("     AA.NATURAL_DATE												\n");
            sbSQL.Append("    ,AA.BF_DAY													\n");
            sbSQL.Append("    ,AA.매출실적													\n");
            sbSQL.Append("    ,AA.누적실적													\n");
            sbSQL.Append("    ,AA.이동계획													\n");
            sbSQL.Append("    ,BB.전월실적													\n");
            sbSQL.Append("    ,AA.N_DT                             		    				\n");
            sbSQL.Append("FROM (        													\n");
            sbSQL.Append("	    SELECT NATURAL_DATE,													\n");
            //sbSQL.Append("           TO_NUMBER(substr(NATURAL_DATE, 5, 2)) || '/' || substr(NATURAL_DATE, 7, 2) AS N_DT,  \n");
            sbSQL.Append("           substr(NATURAL_DATE, 7, 2) || TO_CHAR(TO_DATE(C.NATURAL_DATE), 'dy')   AS N_DT,    \n");
            sbSQL.Append("           SUBSTR(NATURAL_DATE, 7, 2) BF_DAY,                                 \n");
            sbSQL.Append("           매출실적,                                                          \n");
            sbSQL.Append("case when '" + date + "'  = to_char(sysdate, 'yyyymm') and sysdate < NATURAL_DATE then null  \n");
            sbSQL.Append("     ELSE  SUM (매출실적) OVER (ORDER BY NATURAL_DATE) END AS 누적실적,       \n");
            sbSQL.Append("	         SUM(AF.FORCAST_QTY) OVER (ORDER BY NATURAL_DATE) AS 이동계획       \n");


            sbSQL.Append("	    FROM   (SELECT A.REPORT_DATE,                                             \n");
            sbSQL.Append("	                   ROUND(SUM(A.EXCHANGE_SALE + NVL(B.EXCHANGE_SALE,0))/1000000, 0) AS 매출실적     \n");
            sbSQL.Append("	            FROM (                                                          \n");
            sbSQL.Append("	                  SELECT REPORT_DATE                                        \n");
            sbSQL.Append("	                       , SUM(EXCHANGE_SALE) AS EXCHANGE_SALE                \n");
            sbSQL.Append("	                    FROM ASFC_LOT_SALES LC                                  \n");
            sbSQL.Append("	                   WHERE LC.PLANT = 'CCUBEDIGITAL'                          \n");
            sbSQL.Append("	                     AND LC.REPORT_DATE LIKE '" + date + "' || '%'          \n");
            sbSQL.Append("	                     AND LC.CUSTOMER <> 'COMMON'                            \n");
            sbSQL.Append("	                   GROUP BY REPORT_DATE                                     \n");
            sbSQL.Append("	                  )  A,                                                     \n");
            sbSQL.Append("	                  (                                                         \n");
            sbSQL.Append("	                   SELECT SUM(TO_NUMBER(GRP_4))/TO_NUMBER(TO_CHAR(LAST_DAY('" + date + "' || '01'),'DD')) AS EXCHANGE_SALE    \n");
            sbSQL.Append("	                     FROM FRAME_GRPCODEDATA A                               \n");
            sbSQL.Append("	                    WHERE A.PLANT = 'CCUBEDIGITAL'                          \n");
            sbSQL.Append("	                      AND A.GRPTABLE_NAME = 'DEV_SALES'                     \n");
            sbSQL.Append("	                   ) B                                                      \n");
            sbSQL.Append("	                   GROUP BY A.REPORT_DATE                                   \n");
            sbSQL.Append("	            ) DT,                                                           \n");
            sbSQL.Append("	           CALENDAR C,                                                      \n");
            sbSQL.Append("	           ADM_SALE_FORCAST AF                                              \n");
            sbSQL.Append("	    WHERE  C.PLANT = 'CCUBEDIGITAL'                                         \n");
            sbSQL.Append("	    AND    AF.PLANT = C.PLANT                                               \n");
            sbSQL.Append("	    AND    NATURAL_DATE LIKE  '" + date + "' || '%'                          \n");
            sbSQL.Append("	    AND    REPORT_DATE(+) = NATURAL_DATE                                    \n");
            sbSQL.Append("	    AND    NATURAL_DATE = AF.DAY_TIME                                       \n");
            sbSQL.Append("	    GROUP BY NATURAL_DATE, 매출실적, AF.FORCAST_QTY                         \n");
            sbSQL.Append("    )AA                                                                       \n");
            sbSQL.Append("    ,(                                                                        \n");
            sbSQL.Append("        SELECT BF_DAY, SUM(매출실적) OVER (                                    \n");
            sbSQL.Append("                ORDER BY REPORT_DATE) AS 전월실적                              \n");
            sbSQL.Append("        FROM   (SELECT REPORT_DATE,                                                 \n");
            sbSQL.Append("                       SUBSTR(REPORT_DATE, 7, 2) AS BF_DAY,                         \n");
            sbSQL.Append("                       ROUND (SUM (EXCHANGE_SALE) / 1000000, 0) AS 매출실적         \n");
            sbSQL.Append("                FROM   ASFC_LOT_SALES                             \n");
            sbSQL.Append("                WHERE  PLANT ='CCUBEDIGITAL'                      \n");
            sbSQL.Append("                  AND  REPORT_DATE LIKE TO_CHAR(ADD_MONTHS(TO_DATE('" + date + "', 'YYYYMM'), -1), 'YYYYMM') || '%' \n");
            sbSQL.Append("                GROUP BY REPORT_DATE                              \n");
            sbSQL.Append("                ORDER BY REPORT_DATE)                             \n");
            sbSQL.Append("      )BB                                                         \n");
            sbSQL.Append("WHERE AA.BF_DAY = BB.BF_DAY(+)                                    \n");
            sbSQL.Append("ORDER BY 1                                       \n");
            
            return sbSQL.ToString();
        }

        private string getSQL2(string GROUP)
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("			WITH A AS(																			\n");
            sbSQL.Append("			        SELECT CUSTOMER,															\n");
            sbSQL.Append("			               ROUND (NVL (SUM (실적), 0) / 1000000, 2) AS TOT                      \n");
            sbSQL.Append("					FROM   (SELECT SD.SYSCODE_GROUP,                                                \n");
            sbSQL.Append("					               CASE SYSCODE_NAME                                                \n");
            sbSQL.Append("					                  WHEN 'DDI_BUMP' THEN PS.WAFER_DIA || '\"' || SYSCODE_NAME     \n");
            sbSQL.Append("					                  WHEN 'WLP_BUMP' THEN PS.WAFER_DIA || '\"' || SYSCODE_NAME     \n");
            sbSQL.Append("					                  ELSE SYSCODE_NAME                                             \n");
            sbSQL.Append("					                END AS SYSCODE_NAME,                                            \n");
            sbSQL.Append("					               LC.EXCHANGE_SALE AS 실적,                                        \n");
            sbSQL.Append("					               PS.CUSTOMER                                                      \n");
            sbSQL.Append("					        FROM   ASFC_LOT_SALES LC,                                               \n");
            sbSQL.Append("					               PARTSPEC PS,                                                     \n");
            sbSQL.Append("					               SYSCODEDATA SD                                                   \n");
            sbSQL.Append("					        WHERE  LC.PLANT = 'CCUBEDIGITAL'                                        \n");
            sbSQL.Append("					        AND    PS.PLANT = LC.PLANT                                              \n");
            sbSQL.Append("					        AND    SD.PLANT = PS.PLANT                                              \n");
            sbSQL.Append("					        AND    PS.PART_ID = LC.PART                                             \n");
            sbSQL.Append("					        AND    LC.SHOP_CODE = SD.SYSCODE_NAME                                   \n");
            sbSQL.Append("					        AND    SD.SYSTABLE_NAME = 'SHOP_CODE'                                   \n");
            sbSQL.Append("					        AND    SD.SYSCODE_GROUP = '"+GROUP+"'                                   \n");
            sbSQL.Append("					        AND    LC.REPORT_DATE LIKE '" + date + "' ||'%'                         \n");
            //개발품 추가 20160525
            sbSQL.Append("					        UNION ALL                                                              \n");
            sbSQL.Append("					        SELECT GRP_2 AS SYSCODE_GROUP, GRP_5 AS SYSCODE_NAME                   \n");
            sbSQL.Append("					        , SUM(TO_NUMBER(GRP_4))/TO_NUMBER(TO_CHAR(LAST_DAY('" + date + "'||'01'),'DD'))              \n");
            sbSQL.Append("					           *   CASE when '201605' = to_char(sysdate, 'yyyymm') THEN TO_NUMBER(TO_CHAR(SYSDATE,'DD')) \n");
            sbSQL.Append("					                    ELSE TO_NUMBER(TO_CHAR(LAST_DAY('201605'||'01'),'DD'))                           \n");
            sbSQL.Append("					                     END AS EXCHANGE_SALE                                                            \n");
            sbSQL.Append("					        ,GRP_7 AS CUSTOMER                                                       \n");
            sbSQL.Append("					        FROM FRAME_GRPCODEDATA A                                                 \n");
            sbSQL.Append("					        WHERE A.PLANT = 'CCUBEDIGITAL'                                           \n");
            sbSQL.Append("					          AND A.GRPTABLE_NAME = 'DEV_SALES'                                      \n");
            sbSQL.Append("					          AND GRP_2 = '" + GROUP + "'                                            \n");
            sbSQL.Append("					        GROUP BY GRP_2, GRP_5, GRP_7                                             \n");
            //개발품 추가 20160525
            sbSQL.Append("					   )                                                                             \n");
            sbSQL.Append("					GROUP BY (CUSTOMER)                                                             \n");
            sbSQL.Append("			),                                                                                  \n");
            sbSQL.Append("			B AS(                                                                               \n");
            sbSQL.Append("			SELECT 													                            \n");
            sbSQL.Append("			               ROUND (NVL (SUM (실적), 0) / 1000000, 2) AS S_TOT                    \n");
            sbSQL.Append("					FROM   (SELECT SD.SYSCODE_GROUP,                                                \n");
            sbSQL.Append("					               CASE SYSCODE_NAME                                                \n");
            sbSQL.Append("					                  WHEN 'DDI_BUMP' THEN PS.WAFER_DIA || '\"' || SYSCODE_NAME     \n");
            sbSQL.Append("					                  WHEN 'WLP_BUMP' THEN PS.WAFER_DIA || '\"' || SYSCODE_NAME     \n");
            sbSQL.Append("					                  ELSE SYSCODE_NAME                                             \n");
            sbSQL.Append("					                END AS SYSCODE_NAME,                                            \n");
            sbSQL.Append("					               LC.EXCHANGE_SALE AS 실적,                                        \n");
            sbSQL.Append("					               PS.CUSTOMER                                                      \n");
            sbSQL.Append("					        FROM   ASFC_LOT_SALES LC,                                               \n");
            sbSQL.Append("					               PARTSPEC PS,                                                     \n");
            sbSQL.Append("					               SYSCODEDATA SD                                                   \n");
            sbSQL.Append("					        WHERE  LC.PLANT = 'CCUBEDIGITAL'                                        \n");
            sbSQL.Append("					        AND    PS.PLANT = LC.PLANT                                              \n");
            sbSQL.Append("					        AND    SD.PLANT = PS.PLANT                                              \n");
            sbSQL.Append("					        AND    PS.PART_ID = LC.PART                                             \n");
            sbSQL.Append("					        AND    LC.SHOP_CODE = SD.SYSCODE_NAME                                   \n");
            sbSQL.Append("					        AND    SD.SYSTABLE_NAME = 'SHOP_CODE'                                   \n");
            sbSQL.Append("					        AND    SD.SYSCODE_GROUP = '" + GROUP + "'                               \n");
            sbSQL.Append("					        AND    LC.REPORT_DATE LIKE '" + date + "' ||'%'           \n");
            //개발품 추가 20160525
            sbSQL.Append("					        UNION ALL                                                              \n");
            sbSQL.Append("					        SELECT GRP_2 AS SYSCODE_GROUP, GRP_5 AS SYSCODE_NAME                   \n");
            sbSQL.Append("					        , SUM(TO_NUMBER(GRP_4))/TO_NUMBER(TO_CHAR(LAST_DAY('" + date + "'||'01'),'DD'))              \n");
            sbSQL.Append("					           *   CASE when '201605' = to_char(sysdate, 'yyyymm') THEN TO_NUMBER(TO_CHAR(SYSDATE,'DD')) \n");
            sbSQL.Append("					                    ELSE TO_NUMBER(TO_CHAR(LAST_DAY('201605'||'01'),'DD'))                           \n");
            sbSQL.Append("					                     END AS EXCHANGE_SALE                                                            \n");
            sbSQL.Append("					        ,GRP_7 AS CUSTOMER                                                       \n");
            sbSQL.Append("					        FROM FRAME_GRPCODEDATA A                                                 \n");
            sbSQL.Append("					        WHERE A.PLANT = 'CCUBEDIGITAL'                                           \n");
            sbSQL.Append("					          AND A.GRPTABLE_NAME = 'DEV_SALES'                                      \n");
            sbSQL.Append("					          AND GRP_2 = '" + GROUP + "'                                            \n");
            sbSQL.Append("					        GROUP BY GRP_2, GRP_5, GRP_7                                             \n");
            //개발품 추가 20160525
            sbSQL.Append("			)                                                                                   \n");
            sbSQL.Append("			),                                                                                   \n");
            sbSQL.Append("            p as(																			\n");
            sbSQL.Append("                SELECT                                                                    \n");
            sbSQL.Append("                     CUSTOMER                                                             \n");
            sbSQL.Append("                    ,TOT                                                                  \n");
            sbSQL.Append("                    ,S_TOT                                                                \n");
            sbSQL.Append("                    ,CASE WHEN TOT != 0 THEN ROUND((TOT /  S_TOT), 2)                     \n");
            sbSQL.Append("                          ELSE 0                                                          \n");
            sbSQL.Append("                           END AS PROE                                                    \n");
            sbSQL.Append("                FROM A, B                                                                 \n");
            sbSQL.Append("            )                                                                             \n");
            sbSQL.Append("            select CUSTOMER, TOT, S_TOT, PROE, PROE * 100 || '%' as proe1 from p WHERE PROE > 0.01                    \n");
            sbSQL.Append("            union all                                                                     \n");
            sbSQL.Append("            select 'Other', SUM(TOT), MAX(S_TOT), SUM(PROE), SUM(PROE) * 100 || '%' as proe1 from p WHERE PROE <= 0.01     \n");
            return sbSQL.ToString();
        }

        private string getSQL3()
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("			WITH A AS                                                                  	  \n");
            sbSQL.Append("			(       SELECT NATURAL_DATE,                                                  \n");
            sbSQL.Append("			               매출실적,                                                      \n");
            sbSQL.Append("			               SUM (매출실적) OVER (                                          \n");
            sbSQL.Append("			                ORDER BY NATURAL_DATE) AS 누적실적,                           \n");
            sbSQL.Append("			               SUM(AF.FORCAST_QTY) OVER (                                     \n");
            sbSQL.Append("			                ORDER BY NATURAL_DATE) AS 이동계획                            \n");
            sbSQL.Append("	    FROM   (SELECT A.REPORT_DATE,                                             \n");
            sbSQL.Append("	                   ROUND(SUM(A.EXCHANGE_SALE + NVL(B.EXCHANGE_SALE,0))/1000000, 0) AS 매출실적     \n");
            sbSQL.Append("	            FROM (                                                          \n");
            sbSQL.Append("	                  SELECT REPORT_DATE                                        \n");
            sbSQL.Append("	                       , SUM(EXCHANGE_SALE) AS EXCHANGE_SALE                \n");
            sbSQL.Append("	                    FROM ASFC_LOT_SALES LC                                  \n");
            sbSQL.Append("	                   WHERE LC.PLANT = 'CCUBEDIGITAL'                          \n");
            sbSQL.Append("	                     AND LC.REPORT_DATE LIKE '" + date + "' || '%'          \n");
            sbSQL.Append("	                     AND LC.CUSTOMER <> 'COMMON'                            \n");
            sbSQL.Append("	                   GROUP BY REPORT_DATE                                     \n");
            sbSQL.Append("	                  )  A,                                                     \n");
            sbSQL.Append("	                  (                                                         \n");
            sbSQL.Append("	                   SELECT SUM(TO_NUMBER(GRP_4))/TO_NUMBER(TO_CHAR(LAST_DAY('" + date + "' || '01'),'DD')) AS EXCHANGE_SALE    \n");
            sbSQL.Append("	                     FROM FRAME_GRPCODEDATA A                               \n");
            sbSQL.Append("	                    WHERE A.PLANT = 'CCUBEDIGITAL'                          \n");
            sbSQL.Append("	                      AND A.GRPTABLE_NAME = 'DEV_SALES'                     \n");
            sbSQL.Append("	                   ) B                                                      \n");
            sbSQL.Append("	                   GROUP BY A.REPORT_DATE                                   \n");
            sbSQL.Append("	            ) DT,                                                           \n");
            sbSQL.Append("			               CALENDAR C,                                                    \n");
            sbSQL.Append("			               ADM_SALE_FORCAST AF                                            \n");
            sbSQL.Append("			        WHERE  C.PLANT = 'CCUBEDIGITAL'                                       \n");
            sbSQL.Append("			        AND    AF.PLANT = C.PLANT                                             \n");
            sbSQL.Append("			        AND    NATURAL_DATE LIKE '" + date + "' || '%'                              \n");
            sbSQL.Append("			        AND    REPORT_DATE(+) = NATURAL_DATE                                  \n");
            sbSQL.Append("			        AND    NATURAL_DATE = AF.DAY_TIME                                     \n");
            sbSQL.Append("			        GROUP BY NATURAL_DATE, 매출실적, AF.FORCAST_QTY                       \n");
            sbSQL.Append("			),                                                                            \n");
            sbSQL.Append("			B AS(                                                                         \n");
            sbSQL.Append("			    select max(누적실적)  AS TODAY_누적실적,                                  \n");
            sbSQL.Append("			           max(이동계획)  AS MONTH_이동계획,                                  \n");
            sbSQL.Append("			           max(NATURAL_DATE) as NATURAL_DATE                                  \n");
            sbSQL.Append("			    FROM   A                                                                  \n");
            sbSQL.Append("			),                                                                            \n");
            sbSQL.Append("			C AS(                                                                         \n");
            sbSQL.Append("			    SELECT 이동계획 AS TODAY_이동계획                                         \n");
            sbSQL.Append("			    FROM   (                                                                  \n");
            sbSQL.Append("			    SELECT case when '" + date + "' = to_char(sysdate, 'yyyymm') and max(NATURAL_DATE) < to_char(sysdate,'yyyymmdd') then to_char(sysdate-1, 'yyyymmdd') \n");
            sbSQL.Append("			                else MAX(NATURAL_DATE)                           \n");
            sbSQL.Append("			                 end AS NATURAL_DATE                             \n");
            sbSQL.Append("			            FROM   A                                                          \n");
            sbSQL.Append("			            WHERE  NATURAL_DATE <= '" + date + "' || '31'                     \n");
            sbSQL.Append("			            AND    매출실적 IS NOT NULL )DT ,                                 \n");
            sbSQL.Append("			           A                                                                  \n");
            sbSQL.Append("			    WHERE  DT.NATURAL_DATE = A.NATURAL_DATE                                   \n");
            sbSQL.Append("			)                                                                             \n");
            sbSQL.Append("			SELECT                                                                        \n");
            sbSQL.Append("			     ROUND(TODAY_누적실적 / TODAY_이동계획, 2) AS D1                          \n");
            sbSQL.Append("			    ,ROUND(TODAY_누적실적 / MONTH_이동계획, 2) AS D2                          \n");
            sbSQL.Append("			    ,CASE WHEN ROUND(																																																																		\n");
            sbSQL.Append("                                    CASE WHEN '" + date + "' = TO_CHAR(SYSDATE, 'YYYYMM') THEN TO_CHAR(SYSDATE-1, 'DD')                                               \n");
            sbSQL.Append("                                         WHEN '" + date + "' !=  TO_CHAR(SYSDATE, 'YYYYMM') THEN TO_CHAR(ADD_MONTHS(TO_DATE('" + date + "', 'YYYYMM'),+1) - 1, 'DD')      \n");
            sbSQL.Append("                                         END / SUBSTR(NATURAL_DATE, 7, 2), 2                                                                                  \n");
            sbSQL.Append("                                 )  > 1 THEN 1                                                                                                                \n");
            sbSQL.Append("                      ELSE ROUND(                                                                                                                             \n");
            sbSQL.Append("                                    CASE WHEN '" + date + "' = TO_CHAR(SYSDATE, 'YYYYMM') THEN TO_CHAR(SYSDATE-1, 'DD')                                               \n");
            sbSQL.Append("                                         WHEN '" + date + "' !=  TO_CHAR(SYSDATE, 'YYYYMM') THEN TO_CHAR(ADD_MONTHS(TO_DATE('" + date + "', 'YYYYMM'),+1) - 1, 'DD')      \n");
            sbSQL.Append("                                         END / SUBSTR(NATURAL_DATE, 7, 2), 2                                                                                  \n");
            sbSQL.Append("                                 )                                                                                                                            \n");
            sbSQL.Append("                      END  AS D3 , MONTH_이동계획 AS D4,TODAY_누적실적 AS D5                                                                                                                           \n");
            sbSQL.Append("			FROM B, C                                                                     \n");
            return sbSQL.ToString();
        }
    }
}