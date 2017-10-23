using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.MB.MB_B0002
{
    public partial class MB_B0002 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string sql_cust_cd;
        string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        DataSet ds = new DataSet(); 

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //http://192.168.10.98:369/ERPAddition/MB/MB_B0002/MB_B0002.aspx?usermail=janghs0501@nepes.co.kr;
                if (Request.QueryString["usermail"] == null || Request.QueryString["usermail"] == "")
                {
                    TXT_EMAIL.Text = "sonsh0921@NEPES.CO.KR";
                }
                else
                {
                    TXT_EMAIL.Text = Request.QueryString["usermail"];                    
                }

                DateTime frDate = DateTime.Today.AddDays(-7);
                DateTime toDate = DateTime.Today.AddDays(-1);
                tb_fr_yyyymmdd.Text = frDate.Year.ToString("0000") + frDate.Month.ToString("00") + frDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = toDate.Year.ToString("0000") + toDate.Month.ToString("00") + toDate.Day.ToString("00");

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

        protected void Button1_Click(object sender, EventArgs e)
        {
            exe();
        }

        private void exe()
        {
            ReportViewer1.Reset();
            if (TXT_EMAIL.Text == "" || TXT_EMAIL.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"E-mail 주소를 입력해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }
            else
            {
                setGrid();
            }
        }

        private void setGrid()
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

                /*seq 가 a는 필수 항목 없는경우 조회 불가*/

                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    string script = "alert(\"조회된 데이터가 없습니다.\n(E-mail 및 Date 를 확인해주세요).\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_MB_B0002.rdlc");
                ReportViewer1.LocalReport.DisplayName = "개인별 마법노트 작성현황조회"+ TXT_EMAIL.Text;
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";                
                rds.Value = ds.Tables["DataSet1"];
                ReportViewer1.LocalReport.DataSources.Add(rds);

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
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            string strEmail = TXT_EMAIL.Text.Trim();
            /* 실적 조회 쿼리   20151116 박지영 과장 요청 내가 쓴글과 내가 작성한 댓글의 합으로 한가지만 출력하길 원함*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT 																													  \n");
            sbSQL.Append("	 NAME                                                                                                                     \n");
            sbSQL.Append("	,ISNULL(YYYYMMDD, '[  합          계  ]') AS YYYYMMDD                                                                     \n");
            sbSQL.Append("	,EMAIL                                                                                                                    \n");
            sbSQL.Append("	,MY_NOTE + RE_NOTE as  MY_NOTE                                                                                            \n");
            //sbSQL.Append("	,SND_NOTE                                                                                                                 \n");
            //sbSQL.Append("	,RCV_NOTE                                                                                                                 \n");
            //sbSQL.Append("	,RE_NOTE                                                                                                                  \n");
            sbSQL.Append("FROM(                                                                                                                       \n");
            sbSQL.Append("	SELECT 																													  \n");
            sbSQL.Append("		 MAX(NAME) AS NAME	                                                                                                    \n");
            sbSQL.Append("		,SUBSTRING(YYYYMMDD, 1, 4) + '년 '+ SUBSTRING(YYYYMMDD, 5, 2) + '월 ' + SUBSTRING(YYYYMMDD, 7, 2) + '일' AS YYYYMMDD    \n");
            sbSQL.Append("		,MAX(EMAIL) AS EMAIL                                                                                                    \n");
            sbSQL.Append("		,SUM(CONVERT(INT, MY_NOTE)) AS MY_NOTE                                                                                  \n");
            sbSQL.Append("		,SUM(CONVERT(INT,SND_NOTE)) AS SND_NOTE                                                                                 \n");
            sbSQL.Append("		,SUM(CONVERT(INT,RCV_NOTE)) AS RCV_NOTE                                                                                 \n");
            sbSQL.Append("		,SUM(CONVERT(INT,RE_NOTE)) AS RE_NOTE                                                                                   \n");
            sbSQL.Append("	FROM dbo.H_MABUPNOTE                                                                                                      \n");
            sbSQL.Append(" WHERE EMAIL= '" + strEmail + "'                                                                                         \n");
            sbSQL.Append("   AND YYYYMMDD BETWEEN '" + strFrom + "' AND '" + strTo + "'                                                            \n");
            sbSQL.Append("	GROUP BY YYYYMMDD WITH ROLLUP                                                                                             \n");
            sbSQL.Append(")TB                                                                                                                         \n");
            sbSQL.Append("ORDER BY 1 ASC                                                                                                              \n");
            return sbSQL.ToString();
        }
    }
}