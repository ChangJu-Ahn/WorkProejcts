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

namespace ERPAppAddition.ERPAddition.AC.AC_A1002
{
    public partial class ac_a1002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_enc"].ConnectionString);
        string userid;

        SqlCommand sql_cmd = new SqlCommand();
        DataSet ds = new DataSet();
        int value;
        string setSQL = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                WebSiteCount();

                /*로그인 ID가져오기*/
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                setCombo();                
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private string searchSql()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT                   ");
            sbSQL.Append("	 PROJ                  ");
            sbSQL.Append("	,PROJ_FROM             ");
            sbSQL.Append("	,PROJ_TO               ");
            sbSQL.Append("	,COMP_YN               ");
            sbSQL.Append("	,CONT_AMT              ");
            sbSQL.Append("	,CONT_NOTE             ");
            sbSQL.Append("	,ACT_RATE              ");
            sbSQL.Append("	,ACT_NOTE              ");
            sbSQL.Append("	,EMP_RATE              ");
            sbSQL.Append("	,EMP_NOTE              ");
            sbSQL.Append("	,EXE_COST_AMT          ");
            sbSQL.Append("	,MANAGER               ");
            sbSQL.Append("	,REVISION              ");
            sbSQL.Append(" FROM B_PROJ_CTL         ");
            sbSQL.Append("WHERE PROJ = '" + TXT_PROJ_NM.Text + "'   ");
            return sbSQL.ToString();
        }

        private string insertSql()
        {
            string compYn = "";
            if(COMP_YES.Checked == true)
            {
                compYn = "Y";
            }else{
                compYn = "N";
            }

            string user = Session["User"].ToString();            

            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("MERGE                                   ");
            sbSQL.Append("	B_PROJ_CTL A                          ");
            sbSQL.Append("USING(                                  ");
            sbSQL.Append("SELECT                                  ");
            sbSQL.Append("	 '" + TXT_PROJ_NM.Text + "' AS PROJ          ");
            sbSQL.Append("	,'" + TXT_PROJ_FROM.Text + "' AS PROJ_FROM   ");
            sbSQL.Append("	,'" + TXT_PROJ_TO.Text + "' AS PROJ_TO                        ");
            sbSQL.Append("	,'" + compYn + "' AS COMP_YN                        ");
            sbSQL.Append("	,'" + TXT_CONT_AMT.Text + "' AS CONT_AMT                       ");
            sbSQL.Append("	,'" + TXT_CONT_NOTE.Text + "' AS CONT_NOTE                      ");
            sbSQL.Append("	,'" + TXT_ACT_RATE.Text + "' AS ACT_RATE                       ");
            sbSQL.Append("	,'" + TXT_ACT_NOTE.Text + "' AS ACT_NOTE                       ");
            sbSQL.Append("	,'" + TXT_EMP_RATE.Text + "' AS EMP_RATE                       ");
            sbSQL.Append("	,'" + TXT_EMP_NOTE.Text + "' AS EMP_NOTE                       ");
            sbSQL.Append("	,'ac_a1002' AS INSRT_PROG_ID                  ");
            sbSQL.Append("	,'" + user + "' AS INSRT_USER_ID                  ");
            sbSQL.Append("	,getdate() AS INSRT_DT                       ");
            sbSQL.Append("	,'ac_a1002' AS UPDT_PROG_ID                   ");
            sbSQL.Append("	,'" + user + "' AS UPDT_USER_ID                   ");
            sbSQL.Append("	,getdate() AS UPDT_DT                        ");
            sbSQL.Append("	,'" + TXT_EXE_COST_AMT.Text + "' AS EXE_COST_AMT             ");
            sbSQL.Append("	,'" + TXT_MANAGER.Text + "' AS MANAGER                       ");
            sbSQL.Append("	,'" + TXT_REVISION.Text + "' AS REVISION                     ");
            sbSQL.Append(")B                                      ");
            sbSQL.Append("ON A.PROJ = B.PROJ                      ");
            sbSQL.Append("WHEN MATCHED THEN                       ");
            sbSQL.Append("UPDATE SET                              ");
            sbSQL.Append("	 A.PROJ_FROM = B.PROJ_FROM            ");
            sbSQL.Append("	,A.PROJ_TO = B.PROJ_TO                ");
            sbSQL.Append("	,A.COMP_YN = B.COMP_YN                ");
            sbSQL.Append("	,A.CONT_AMT = B.CONT_AMT              ");
            sbSQL.Append("	,A.CONT_NOTE = B.CONT_NOTE            ");
            sbSQL.Append("	,A.ACT_RATE = B.ACT_RATE              ");
            sbSQL.Append("	,A.ACT_NOTE = B.ACT_NOTE              ");
            sbSQL.Append("	,A.EMP_RATE = B.EMP_RATE              ");
            sbSQL.Append("	,A.EMP_NOTE = B.EMP_NOTE              ");
            sbSQL.Append("	,A.UPDT_PROG_ID = B.UPDT_PROG_ID              ");
            sbSQL.Append("	,A.UPDT_USER_ID = B.UPDT_USER_ID              ");
            sbSQL.Append("	,A.UPDT_DT = B.UPDT_DT              ");
            sbSQL.Append("	,A.EXE_COST_AMT = B.EXE_COST_AMT              ");
            sbSQL.Append("	,A.MANAGER = B.MANAGER              ");
            sbSQL.Append("	,A.REVISION = B.REVISION              ");
            sbSQL.Append("WHEN NOT MATCHED THEN                   ");
            sbSQL.Append("INSERT (                                ");
            sbSQL.Append("		 PROJ                               ");
            sbSQL.Append("		,PROJ_FROM                          ");
            sbSQL.Append("		,PROJ_TO                            ");
            sbSQL.Append("		,COMP_YN                            ");
            sbSQL.Append("		,CONT_AMT                           ");
            sbSQL.Append("		,CONT_NOTE                          ");
            sbSQL.Append("		,ACT_RATE                           ");
            sbSQL.Append("		,ACT_NOTE                           ");
            sbSQL.Append("		,EMP_RATE                           ");
            sbSQL.Append("		,EMP_NOTE                           ");
            sbSQL.Append("		,INSRT_PROG_ID                      ");
            sbSQL.Append("		,INSRT_USER_ID                      ");
            sbSQL.Append("		,INSRT_DT                           ");
            sbSQL.Append("		,UPDT_PROG_ID                       ");
            sbSQL.Append("		,UPDT_USER_ID                       ");
            sbSQL.Append("		,UPDT_DT                            ");
            sbSQL.Append("		,EXE_COST_AMT                       ");
            sbSQL.Append("		,MANAGER                            ");
            sbSQL.Append("		,REVISION                           ");
            sbSQL.Append(")VALUES(                                  ");
            sbSQL.Append("		 B.PROJ                             ");
            sbSQL.Append("		,B.PROJ_FROM                        ");
            sbSQL.Append("		,B.PROJ_TO                          ");
            sbSQL.Append("		,B.COMP_YN                          ");
            sbSQL.Append("		,B.CONT_AMT                         ");
            sbSQL.Append("		,B.CONT_NOTE                        ");
            sbSQL.Append("		,B.ACT_RATE                         ");
            sbSQL.Append("		,B.ACT_NOTE                         ");
            sbSQL.Append("		,B.EMP_RATE                         ");
            sbSQL.Append("		,B.EMP_NOTE                         ");
            sbSQL.Append("		,B.INSRT_PROG_ID                    ");
            sbSQL.Append("		,B.INSRT_USER_ID                    ");
            sbSQL.Append("		,B.INSRT_DT                         ");
            sbSQL.Append("		,B.UPDT_PROG_ID                     ");
            sbSQL.Append("		,B.UPDT_USER_ID                     ");
            sbSQL.Append("		,B.UPDT_DT                          ");
            sbSQL.Append("		,B.EXE_COST_AMT                     ");
            sbSQL.Append("		,B.MANAGER                          ");
            sbSQL.Append("		,B.REVISION                         ");
            sbSQL.Append("		);                                  ");
            return sbSQL.ToString();
        }


        protected void bt_search_Click(object sender, EventArgs e)
        {
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            sql_cmd = conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = searchSql();
            sql_cmd.CommandTimeout = 3000; 
            
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                if (dt.Rows.Count < 1)
                {
                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);

                    TXT_PROJ_FROM.Text = "";
                    TXT_PROJ_TO.Text = "";
                    TXT_CONT_AMT.Text = "";
                    TXT_CONT_NOTE.Text = "";
                    TXT_ACT_RATE.Text = "";
                    TXT_ACT_NOTE.Text = "";
                    TXT_EMP_RATE.Text = "";
                    TXT_EMP_NOTE.Text = "";
                    TXT_EXE_COST_AMT.Text = "";
                    TXT_MANAGER.Text = "";
                    TXT_REVISION.Text = "";
                    
                    return;
                }
                else
                {
                    /*select 값 set*/
                    TXT_PROJ_FROM.Text = dt.Rows[0]["PROJ_FROM"].ToString();
                    TXT_PROJ_TO.Text = dt.Rows[0]["PROJ_TO"].ToString();
                    if (dt.Rows[0]["COMP_YN"].ToString() == "Y")
                    {
                        COMP_YES.Checked = true;
                    }else
                    {
                        COMP_N0.Checked = true;
                    }
                    TXT_CONT_AMT.Text = dt.Rows[0]["CONT_AMT"].ToString();
                    TXT_CONT_NOTE.Text = dt.Rows[0]["CONT_NOTE"].ToString();
                    TXT_ACT_RATE.Text = dt.Rows[0]["ACT_RATE"].ToString();
                    TXT_ACT_NOTE.Text = dt.Rows[0]["ACT_NOTE"].ToString();
                    TXT_EMP_RATE.Text = dt.Rows[0]["EMP_RATE"].ToString();
                    TXT_EMP_NOTE.Text = dt.Rows[0]["EMP_NOTE"].ToString();
                    TXT_EXE_COST_AMT.Text = dt.Rows[0]["EXE_COST_AMT"].ToString();
                    TXT_MANAGER.Text = dt.Rows[0]["MANAGER"].ToString();
                    TXT_REVISION.Text = dt.Rows[0]["REVISION"].ToString();
                }
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                Console.WriteLine("{0} Second exception caught.", ex);
            }
        }

        public void setCombo()
        {
            /*수량단위 */
            DataTable UNIT = getData("SELECT MINOR_CD, '[' + MINOR_CD + ']' + MINOR_NM as MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'A9003' order by 1");
            if (UNIT.Rows.Count > 0)
            {
                DataRow dr = UNIT.NewRow();
                UNIT.Rows.InsertAt(dr, 0);

                DDL_PROJ.DataTextField = "MINOR_NM";
                DDL_PROJ.DataValueField = "MINOR_CD";
                DDL_PROJ.DataSource = UNIT;
                DDL_PROJ.DataBind();
            }
        }

        public DataTable getData(string sql)
        {
            SqlDataReader sql_dr;
            DataSet ds = new DataSet();

            DataTable retDt = new DataTable();

            try
            {
                // 프로시져 실행: 기본데이타 생성
                conn.Open();
                sql_cmd = conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = sql;

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds);                    
                }
                catch (Exception ex)
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
                conn.Close();

                retDt = ds.Tables[0];

            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return retDt;
        }    


        protected void bt_save_Click(object sender, EventArgs e)
        {
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            sql_cmd = conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = insertSql();
            sql_cmd.CommandTimeout = 3000;

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                string script = "alert(\"저장되었습니다..\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);

                /*재조회*/
                bt_search_Click(sender, e);
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                Console.WriteLine("{0} Second exception caught.", ex);
            }


        }

        protected void DDL_PROJ_SelectedIndexChanged(object sender, EventArgs e)
        {
            TXT_PROJ_NM.Text = DDL_PROJ.SelectedValue.ToString();
        }
    }
}