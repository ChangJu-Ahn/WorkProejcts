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



namespace ERPAppAddition.ERPAddition.SM.sm_s9001
{
    public partial class ds_sm_s9001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];        

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        SqlCommand sql_cmd2 = new SqlCommand();                

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;

        DataSet ds = new DataSet();

        System.DateTime dateTime = System.DateTime.Now.AddDays(-1);

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {                
                setWeek();
                WebSiteCount();
            }

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void setWeek()
        {
            conn.Open();
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select SUBSTR(PLAN_YEAR,3,2)|| '년' || WEEK || '주차(' || NATURAL_DATE ||')'as m_week , SUBSTR(PLAN_YEAR,3,2) || WEEK as week  		\n");
                sbSQL.Append("FROM (                                                                                                                      \n");
                sbSQL.Append("    select PLAN_YEAR, NATURAL_DATE, to_char(to_date(NATURAL_DATE,'yyyymmdd'), 'dy') AS DY, LPAD(PLAN_WEEK,2,'0') AS WEEK    \n");
                sbSQL.Append("      from CALENDAR                                                                                                         \n");
                sbSQL.Append("     where PLANT = 'CCUBEDIGITAL'                                                                                           \n");
                sbSQL.Append("       and PLAN_YEAR >= '2015'                                                                                               \n");
                sbSQL.Append("       and PLAN_YEAR <= to_char(sysdate+8, 'yyyy')                                                                            \n");
                sbSQL.Append("       and NATURAL_DATE <= to_char(sysdate+8,'yyyymmdd')                                                                      \n");
                sbSQL.Append("       )                                                                                                                    \n");
                sbSQL.Append("WHERE DY = '금'                                                                                                             \n");
                sbSQL.Append("order by NATURAL_DATE  desc                                                                                                     \n");

                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                dr = cmd2.ExecuteReader();

                if (dr.RowSize > 0)
                {
                    str_week.DataSource = dr;
                    str_week.DataValueField = "week";
                    str_week.DataTextField = "m_week";
                    str_week.DataBind();
                    str_week.SelectedIndex = 1;
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }            
        }      

        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            if (str_week.Text == "" || str_week.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"Date 를 입력해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }

            else
            {
                setGrid();
            }
        }

        private void setGrid(){
            try
                {
                    // 프로시져 실행: 기본데이타 생성
                    sql_conn.Open();
                    sql_cmd = sql_conn.CreateCommand();
                    sql_cmd.CommandType = CommandType.Text;
                    sql_cmd.CommandText = getSQL();
                    
                    sql_cmd2 = sql_conn.CreateCommand();
                    sql_cmd2.CommandType = CommandType.Text;
                    sql_cmd2.CommandText = getSQL_G();  

                    DataTable dt = new DataTable();
                    try
                    {
                        SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                        da.Fill(ds, "DataSet1");

                        SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                        da2.Fill(ds, "DataSet2");   
                    }
                    catch (Exception ex)
                    {
                        if (sql_conn.State == ConnectionState.Open)
                            sql_conn.Close();
                    }
                    sql_conn.Close();
                    /*seq 가 a는 필수 항목 없는경우 조회 불가*/
                    DataRow[] dr = ds.Tables["DataSet1"].Select("SEQ = 'A'");
                    if(ds.Tables["DataSet1"].Rows.Count <= 0 || dr.Length <= 0)
                    {
                        string script = "alert(\" ["+ str_week.Text + "] 주차 조회된 데이터가 없습니다..\");";
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                        return;
                    }

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s9001.rdlc");
                    ReportViewer1.LocalReport.DisplayName = str_week.Text + "_주차별 전사재고" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = ds.Tables["DataSet1"];
                    ReportViewer1.LocalReport.DataSources.Add(rds);

                    /*그래프*/
                    ReportDataSource rds2 = new ReportDataSource();
                    rds2.Name = "DataSet2";
                    rds2.Value = ds.Tables["DataSet2"];
                    ReportViewer1.LocalReport.DataSources.Add(rds2);


                    ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x                    

                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
        }

        private string getSQL()
        {
            string date = str_week.Text;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT 																																																																																																																													\n");
            sbSQL.Append("	 TT.SEQ, TT.DIVI                                                                                                                                                                                                                                                         \n");
            sbSQL.Append("	,TT.M_STOCK                                                                                                                                                                                                                                                      \n");
            sbSQL.Append("	,LINE_STOCK                                                                                                                                                                                                                                                   \n");
            sbSQL.Append("	,WIP_STOCK                                                                                                                                                                                                                                                    \n");
            sbSQL.Append("	,FGS_STOCK                                                                                                                                                                                                                                                    \n");
            sbSQL.Append("	,GOODS_STOCK                                                                                                                                                                                                                                                  \n");
            sbSQL.Append("	,STOCK_SUM                                                                                                                                                                                                                                                    \n");
            sbSQL.Append("	,FF.M_STOCK AS CYC                                                                                                                                                                                                                                                    \n");
            sbSQL.Append("FROM                                                                                                                                                                                                                                                            \n");
            sbSQL.Append("(                                                                                                                                                                                                                                                               \n");
            sbSQL.Append("		select 'A' AS SEQ                                                                                                                                                                                                                                           \n");
            sbSQL.Append("			,WEEK                                                                                                                                                                                                                                                     \n");
            sbSQL.Append("			,DIVI                                                                                                                                                                                                                                                     \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(M_STOCK/100000000, 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(M_STOCK/100000000, 1)))+1)  AS M_STOCK                                                                                                                 \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(LINE_STOCK/100000000, 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(LINE_STOCK/100000000, 1)))+1)  AS LINE_STOCK                                                                                                        \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(WIP_STOCK/100000000, 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(WIP_STOCK/100000000, 1)))+1)  AS WIP_STOCK                                                                                                           \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(FGS_STOCK/100000000, 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(FGS_STOCK/100000000, 1)))+1)  AS FGS_STOCK                                                                                                           \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(GOODS_STOCK/100000000, 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(GOODS_STOCK/100000000, 1)))+1)  AS GOODS_STOCK                                                                                                     \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, ROUND(M_STOCK/100000000, 1) + ROUND(LINE_STOCK/100000000, 1) + ROUND(WIP_STOCK/100000000, 1) + ROUND(FGS_STOCK/100000000, 1) + ROUND(GOODS_STOCK/100000000, 1)), 1                                                            \n");
            sbSQL.Append("			         , CHARINDEX('.', CONVERT(VARCHAR, ROUND(M_STOCK/100000000, 1) + ROUND(LINE_STOCK/100000000, 1) + ROUND(WIP_STOCK/100000000, 1) + ROUND(FGS_STOCK/100000000, 1) + ROUND(GOODS_STOCK/100000000, 1)                                                 \n");
            sbSQL.Append("			           ))+1)  AS STOCK_SUM			                                                                                                                                                                                                                      \n");
            sbSQL.Append("		from T_TOTAL_STOCK                                                                                                                                                                                                                                          \n");
            sbSQL.Append("		where WEEK = '"+date+"'   \n");
            sbSQL.Append("		 AND GUBN = 'R'           \n");
            sbSQL.Append("		 union all           \n");
            sbSQL.Append("		select 'A' AS SEQ                                                                                                                                                                                                                                           \n");
            sbSQL.Append("			,MAX(WEEK) AS WEEK    \n");
            sbSQL.Append("			,'전사' AS DIVI       \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(M_STOCK/100000000, 1))), 1, CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(M_STOCK/100000000, 1))))+1)  AS M_STOCK                                                                                                                 \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(LINE_STOCK/100000000, 1))), 1, CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(LINE_STOCK/100000000, 1))))+1)  AS LINE_STOCK                                                                                                        \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(WIP_STOCK/100000000, 1))), 1, CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(WIP_STOCK/100000000, 1))))+1)  AS WIP_STOCK                                                                                                           \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(FGS_STOCK/100000000, 1))), 1, CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(FGS_STOCK/100000000, 1))))+1)  AS FGS_STOCK                                                                                                           \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(GOODS_STOCK/100000000, 1))), 1, CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(GOODS_STOCK/100000000, 1))))+1)  AS GOODS_STOCK                                                                                                     \n");
            sbSQL.Append("			,SUBSTRING(CONVERT(VARCHAR, sum(ROUND(M_STOCK/100000000, 1)) + sum(ROUND(LINE_STOCK/100000000, 1)) + sum(ROUND(WIP_STOCK/100000000, 1)) + sum(ROUND(FGS_STOCK/100000000, 1)) + sum(ROUND(GOODS_STOCK/100000000, 1))), 1                                                            \n");
            sbSQL.Append("			         , CHARINDEX('.', CONVERT(VARCHAR, sum(ROUND(M_STOCK/100000000, 1)) + sum(ROUND(LINE_STOCK/100000000, 1)) + sum(ROUND(WIP_STOCK/100000000, 1)) + sum(ROUND(FGS_STOCK/100000000, 1)) + sum(ROUND(GOODS_STOCK/100000000, 1))                                                 \n");
            sbSQL.Append("			           ))+1)  AS STOCK_SUM			                                                                                                                                                                                                                      \n");
            sbSQL.Append("		from T_TOTAL_STOCK                                                                                                                                                                                                                                          \n");
            sbSQL.Append("		where WEEK = '" + date + "'   \n");
            sbSQL.Append("		 AND GUBN = 'R'           \n");
            sbSQL.Append("      GROUP BY WEEK                                                                                                                                                                                                                                                            \n");
            sbSQL.Append("		UNION ALL                                                                                                                                                                                                                                                   \n");
            sbSQL.Append("                                                                                                                                                                                                                                                                \n");
            sbSQL.Append("		SELECT                                                                                                                                                                                                                                                      \n");
            sbSQL.Append("			 SEQ                                                                                                                                                                                                                                                      \n");
            sbSQL.Append("			,WEEK                                                                                                                                                                                                                                                     \n");
            sbSQL.Append("			,DIVI                                                                                                                                                                                                                                                     \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(M_STOCK, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                      \n");
            sbSQL.Append("				  WHEN SUBSTRING(M_STOCK, 1, 1) = '-' THEN '▼' + SUBSTRING(M_STOCK, 2, LEN(M_STOCK))                                                                                                                                                                   \n");
            sbSQL.Append("				  ELSE '▲' + M_STOCK END AS M_STOCK                                                                                                                                                                                                                    \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(LINE_STOCK, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                   \n");
            sbSQL.Append("				  WHEN SUBSTRING(LINE_STOCK, 1, 1) = '-' THEN '▼' + SUBSTRING(LINE_STOCK, 2, LEN(LINE_STOCK))                                                                                                                                                          \n");
            sbSQL.Append("				  ELSE '▲' + LINE_STOCK END AS LINE_STOCK                                                                                                                                                                                                              \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(WIP_STOCK, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                    \n");
            sbSQL.Append("				  WHEN SUBSTRING(WIP_STOCK, 1, 1) = '-' THEN '▼' + SUBSTRING(WIP_STOCK, 2, LEN(WIP_STOCK))                                                                                                                                                             \n");
            sbSQL.Append("				  ELSE '▲' + WIP_STOCK END AS WIP_STOCK                                                                                                                                                                                                                \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(FGS_STOCK, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                    \n");
            sbSQL.Append("				  WHEN SUBSTRING(FGS_STOCK, 1, 1) = '-' THEN '▼' + SUBSTRING(FGS_STOCK, 2, LEN(FGS_STOCK))                                                                                                                                                             \n");
            sbSQL.Append("				  ELSE '▲' + FGS_STOCK END AS FGS_STOCK                                                                                                                                                                                                                \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(GOODS_STOCK, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                  \n");
            sbSQL.Append("				  WHEN SUBSTRING(GOODS_STOCK, 1, 1) = '-' THEN '▼' + SUBSTRING(GOODS_STOCK, 2, LEN(GOODS_STOCK))                                                                                                                                                       \n");
            sbSQL.Append("				  ELSE '▲' + GOODS_STOCK END AS GOODS_STOCK                                                                                                                                                                                                            \n");
            sbSQL.Append("			,CASE WHEN SUBSTRING(STOCK_SUM, 1, 3) = '0.0' THEN '-'                                                                                                                                                                                                    \n");
            sbSQL.Append("				  WHEN SUBSTRING(STOCK_SUM, 1, 1) = '-' THEN '▼' + SUBSTRING(STOCK_SUM, 2, LEN(STOCK_SUM))                                                                                                                                                             \n");
            sbSQL.Append("				  ELSE '▲' + STOCK_SUM END AS STOCK_SUM                                                                                                                                                                                                                \n");
            sbSQL.Append("		FROM(                                                                                                                                                                                                                                                       \n");
            sbSQL.Append("			SELECT 'B' AS SEQ                                                                                                                                                                                                                                         \n");
            sbSQL.Append("				,MAX(WEEK) AS WEEK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				,DIVI                                                                                                                                                                                                                                                   \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1)))+1)  AS M_STOCK                                                                                                     \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(LINE_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(LINE_STOCK), 1)))+1)  AS LINE_STOCK                                                                                            \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(WIP_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(WIP_STOCK), 1)))+1)  AS WIP_STOCK                                                                                               \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(FGS_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(FGS_STOCK), 1)))+1)  AS FGS_STOCK                                                                                               \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(GOODS_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(GOODS_STOCK), 1)))+1)  AS GOODS_STOCK                                                                                         \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1) + ROUND(SUM(LINE_STOCK), 1) + ROUND(SUM(WIP_STOCK), 1) + ROUND(SUM(FGS_STOCK), 1) + ROUND(SUM(GOODS_STOCK), 1)), 1                                 \n");
            sbSQL.Append("				         , CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1) + ROUND(SUM(LINE_STOCK), 1) + ROUND(SUM(WIP_STOCK), 1) + ROUND(SUM(FGS_STOCK), 1) + ROUND(SUM(GOODS_STOCK), 1)                      \n");
            sbSQL.Append("				           ))+1)  AS STOCK_SUM                     				                                                                                                                                                                                              \n");
            sbSQL.Append("			FROM                                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				(                                                                                                                                                                                                                                                       \n");
            sbSQL.Append("				select WEEK, DIVI, round(M_STOCK/100000000, 1) as M_STOCK, ROUND(LINE_STOCK/100000000, 1) as LINE_STOCK, ROUND(WIP_STOCK/100000000, 1) as WIP_STOCK, ROUND(FGS_STOCK/100000000, 1) as FGS_STOCK, ROUND(GOODS_STOCK/100000000, 1) as GOODS_STOCK         \n");
            sbSQL.Append("				from T_TOTAL_STOCK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("			    where WEEK = '"+date+"'\n");
            sbSQL.Append("		          AND GUBN = 'R'       \n");
            sbSQL.Append("				UNION ALL                                                                                                                                                                                                                                               \n");
            sbSQL.Append("				select WEEK, DIVI, ROUND((M_STOCK * -1) / 100000000, 1) as M_STOCK, ROUND((LINE_STOCK * -1) / 100000000, 1) as LINE_STOCK , ROUND((WIP_STOCK * -1) / 100000000, 1) as WIP_STOCK, ROUND((FGS_STOCK * -1) / 100000000, 1) as FGS_STOCK, ROUND((GOODS_STOCK * -1) /100000000, 1) as GOODS_STOCK  \n");
            sbSQL.Append("				from T_TOTAL_STOCK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				where WEEK = -- '"+date+"' -1 --지난주정보                                                                                                                                                                                                              \n");
            sbSQL.Append("				            (SELECT MAX(WEEK) FROM T_TOTAL_STOCK    \n");
            sbSQL.Append("							  WHERE WEEK < '" + date + "')   --전주 정보 구하기   \n");
            sbSQL.Append("		          AND GUBN = 'R'       \n");
            sbSQL.Append("  			) CC                                                                                                                                                                                                                                                    \n");
            sbSQL.Append("			GROUP BY DIVI                                                                                                                                                                                                                                             \n");
            sbSQL.Append("			union all                                                                                                                                                                                                                                            \n");
            sbSQL.Append("			SELECT 'B' AS SEQ                                                                                                                                                                                                                                         \n");
            sbSQL.Append("				,'" + date + "' AS WEEK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				,'전사' as DIVI                                                                                                                                                                                                                                                   \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1)))+1)  AS M_STOCK                                                                                                     \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(LINE_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(LINE_STOCK), 1)))+1)  AS LINE_STOCK                                                                                            \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(WIP_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(WIP_STOCK), 1)))+1)  AS WIP_STOCK                                                                                               \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(FGS_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(FGS_STOCK), 1)))+1)  AS FGS_STOCK                                                                                               \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(GOODS_STOCK), 1)), 1, CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(GOODS_STOCK), 1)))+1)  AS GOODS_STOCK                                                                                         \n");
            sbSQL.Append("				,SUBSTRING(CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1) + ROUND(SUM(LINE_STOCK), 1) + ROUND(SUM(WIP_STOCK), 1) + ROUND(SUM(FGS_STOCK), 1) + ROUND(SUM(GOODS_STOCK), 1)), 1                                 \n");
            sbSQL.Append("				         , CHARINDEX('.', CONVERT(VARCHAR, ROUND(SUM(M_STOCK), 1) + ROUND(SUM(LINE_STOCK), 1) + ROUND(SUM(WIP_STOCK), 1) + ROUND(SUM(FGS_STOCK), 1) + ROUND(SUM(GOODS_STOCK), 1)                      \n");
            sbSQL.Append("				           ))+1)  AS STOCK_SUM                     				                                                                                                                                                                                              \n");
            sbSQL.Append("			FROM                                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				(                                                                                                                                                                                                                                                       \n");
            sbSQL.Append("				select WEEK, DIVI, round(M_STOCK/100000000, 1) as M_STOCK, ROUND(LINE_STOCK/100000000, 1) as LINE_STOCK, ROUND(WIP_STOCK/100000000, 1) as WIP_STOCK, ROUND(FGS_STOCK/100000000, 1) as FGS_STOCK, ROUND(GOODS_STOCK/100000000, 1) as GOODS_STOCK         \n");
            sbSQL.Append("				from T_TOTAL_STOCK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("			    where WEEK = '" + date + "'\n");
            sbSQL.Append("		          AND GUBN = 'R'       \n");
            sbSQL.Append("				UNION ALL                                                                                                                                                                                                                                               \n");
            sbSQL.Append("				select WEEK, DIVI, ROUND((M_STOCK * -1) / 100000000, 1) as M_STOCK, ROUND((LINE_STOCK * -1) / 100000000, 1) as LINE_STOCK , ROUND((WIP_STOCK * -1) / 100000000, 1) as WIP_STOCK, ROUND((FGS_STOCK * -1) / 100000000, 1) as FGS_STOCK, ROUND((GOODS_STOCK * -1) /100000000, 1) as GOODS_STOCK  \n");
            sbSQL.Append("				from T_TOTAL_STOCK                                                                                                                                                                                                                                      \n");
            sbSQL.Append("				where WEEK = -- '" + date + "' -1 --지난주정보                                                                                                                                                                                                              \n");
            sbSQL.Append("				            (SELECT MAX(WEEK) FROM T_TOTAL_STOCK    \n");
            sbSQL.Append("							  WHERE WEEK < '" + date + "')   --전주 정보 구하기   \n");
            sbSQL.Append("		          AND GUBN = 'R'           \n");
            sbSQL.Append("  			) CC                                                                                                                                                                                                                                                 \n");            
            sbSQL.Append("		) TB                                                                                                                                                                                                                                                        \n");
            sbSQL.Append(") TT                                                                                                                                                                                                                                                            \n");
            sbSQL.Append("left join(																																											\n");
            sbSQL.Append("			select WEEK, DIVI, ROUND(M_STOCK, 0)AS M_STOCK, 'A' AS SEQ                                \n");
            sbSQL.Append("			from T_TOTAL_STOCK                                                                        \n");
            sbSQL.Append("			where GUBN = 'G'                                                                          \n");
            sbSQL.Append("			  and WEEK = '" + date + "'                                                                        \n");
            sbSQL.Append("			 UNION ALL                                                                                \n");
            sbSQL.Append("			select WEEK, '전사' AS DIVI, SUM(ROUND(M_STOCK, 0)) AS M_STOCK, 'A' AS SEQ           \n");
            sbSQL.Append("			from T_TOTAL_STOCK                                                                        \n");
            sbSQL.Append("			where GUBN = 'G'                                                                          \n");
            sbSQL.Append("			  and WEEK = '" + date + "'            \n");
            sbSQL.Append("			 GROUP BY WEEK                                                                            \n");
            sbSQL.Append("			    ) FF                                                                                  \n");
            sbSQL.Append("on TT.DIVI = FF.DIVI                                                                            \n");
            //sbSQL.Append("AND TT.SEQ = FF.SEQ                                                                             \n");
            sbSQL.Append("ORDER BY TT.DIVI DESC, TT.SEQ                                                                   \n");
            return sbSQL.ToString();
        }

        private string getSQL_G()
        {
            string date = str_week.Text;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select 'B' AS SEQ																																																							\n");
            sbSQL.Append("      , ST.WEEK AS WEEK                                                                                                       \n");
            sbSQL.Append("      ,ST.DIVI                                                                                                                \n");
            sbSQL.Append("      ,ST.M_STOCK                                                                                                             \n");
            sbSQL.Append("from T_TOTAL_STOCK ST                                                                                                         \n");
            sbSQL.Append("	,(                                                                                                                          \n");
            sbSQL.Append("		select week                                                                                                               \n");
            sbSQL.Append("		from(                                                                                                                     \n");
            sbSQL.Append("		 select ROW_NUMBER() over(order by week desc) as seq                                                                      \n");
            sbSQL.Append("			  ,week                                                                                                                 \n");
            sbSQL.Append("		from T_TOTAL_STOCK                                                                                                        \n");
            sbSQL.Append("		where GUBN = 'G'                                                                                                          \n");
            sbSQL.Append("		  and WEEK <= '" + date + "'                                                                                              \n");
            sbSQL.Append("		group by WEEK                                                                                                             \n");
            sbSQL.Append("		)tb                                                                                                                       \n");
            sbSQL.Append("		WHERE seq BETWEEN 1 AND 4                                                                                                 \n");
            sbSQL.Append("	)ST2                                                                                                                        \n");
            sbSQL.Append("where ST.GUBN = 'G'                                                                                                           \n");
            sbSQL.Append("  AND ST.WEEK = ST2.WEEK                                                                                                      \n");
            sbSQL.Append("                                                                                                                              \n");
            sbSQL.Append("UNION ALL                                                                                                                     \n");
            sbSQL.Append("                                                                                                                              \n");
            sbSQL.Append("select 'A' AS SEQ                                                                                                             \n");
            sbSQL.Append("      ,ST.YYYYMM AS WEEK                                                                                                      \n");
            sbSQL.Append("      ,ST.DIVI                                                                                                                \n");
            sbSQL.Append("      ,AVG(ST.M_STOCK) AS M_STOCK                                                                                             \n");
            sbSQL.Append("from T_TOTAL_STOCK ST                                                                                                         \n");
            sbSQL.Append("	,(                                                                                                                          \n");
            sbSQL.Append("		select YYYYMM                                                                                                             \n");
            sbSQL.Append("		from(                                                                                                                     \n");
            sbSQL.Append("			 select ROW_NUMBER() over(order by YYYYMM desc) as seq                                                                  \n");
            sbSQL.Append("				  ,YYYYMM                                                                                                             \n");
            sbSQL.Append("			from T_TOTAL_STOCK                                                                                                      \n");
            sbSQL.Append("			where GUBN = 'G'                                                                                                        \n");
            sbSQL.Append("			  and WEEK <= '" + date + "'                                                                                            \n");
            sbSQL.Append("			  AND YYYYMM <> (SELECT YYYYMM FROM T_TOTAL_STOCK WHERE WEEK = '" + date + "'  AND GUBN = 'G' GROUP BY YYYYMM )        \n");
            sbSQL.Append("			group by YYYYMM                                                                                                         \n");
            sbSQL.Append(" 		)tb                                                                                                                       \n");
            sbSQL.Append(" 		WHERE SEQ BETWEEN 1 AND 2                                                                                                 \n");
            sbSQL.Append("	)ST2                                                                                                                        \n");
            sbSQL.Append("where ST.GUBN = 'G'                                                                                                           \n");
            sbSQL.Append("  AND ST.YYYYMM = ST2.YYYYMM                                                                                                  \n");
            sbSQL.Append("GROUP BY ST.YYYYMM, ST.DIVI                                                                                                   \n");
            sbSQL.Append("                                                                                                                              \n");
            sbSQL.Append("ORDER BY SEQ, WEEK, DIVI DESC                                                                                                 \n");
            return sbSQL.ToString();
        }
    }
}