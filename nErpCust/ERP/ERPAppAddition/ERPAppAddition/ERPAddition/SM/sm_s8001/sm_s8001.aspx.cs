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



namespace ERPAppAddition.ERPAddition.SM.sm_s8001
{
    public partial class sm_s8001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];        

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();                
        DataSet ds = new DataSet();

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;
        

        System.DateTime dateTime = System.DateTime.Now.AddDays(-1);

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {                
                str_date.Text = dateTime.Year.ToString("00") + dateTime.Month.ToString("00") + dateTime.Day.ToString("00");
                /*달력셋*/
                setMonth();
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

        private void setMonth()
        {
            conn.Open();
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select PLAN_YEAR || lpad(PLAN_MONTH, 2, 0) AS M_MONTH   \n");
                sbSQL.Append("  from CALENDAR                                         \n");
                sbSQL.Append(" where PLANT = 'CCUBEDIGITAL'                           \n");
                sbSQL.Append("   and PLAN_YEAR >= '2015'                              \n");
                sbSQL.Append("   and PLAN_YEAR <= to_char(sysdate, 'yyyy')            \n");
                sbSQL.Append("   and LPAD(PLAN_MONTH, 2, 0) <= to_char(sysdate, 'mm') \n");
                sbSQL.Append(" group by PLAN_YEAR, PLAN_MONTH                         \n");
                sbSQL.Append(" order by 1 desc                                            \n");
                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                dr = cmd2.ExecuteReader();

                if (dr.RowSize > 0)
                {
                    tb_yyyymm.DataSource = dr;
                    tb_yyyymm.DataValueField = "M_MONTH";
                    tb_yyyymm.DataTextField = "M_MONTH";
                    tb_yyyymm.DataBind();
                    tb_yyyymm.SelectedValue = dateTime.Year.ToString() + dateTime.Month.ToString("00");
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
            if (str_date.Text == "" || str_date.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"조회년도를 입력해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }

            else
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
                    if(ds.Tables["DataSet1"].Rows.Count <= 0)
                    {
                        string script = "alert(\"조회된 데이터가 없습니다..\");";
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                        return;
                    }

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s8001.rdlc");
                    ReportViewer1.LocalReport.DisplayName = str_date.Text + "_비가동 Loss 상각비" + DateTime.Now.ToShortDateString();
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
        }

        private string getSQL()
        {            
            
            string date = "";
            if(chk_date.Checked == true)
            {
                date = str_date.Text;
            }
            else
            {
                date = tb_yyyymm.Text + "%";
            }
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_NON_OPER_LOSS  '" + date + "'\n");
            return sbSQL.ToString();
        }

        protected void chk_date_CheckedChanged(object sender, EventArgs e)
        {
            if(chk_date.Checked = true)
            {
                tb_yyyymm.Enabled = false;
                str_date.Enabled = true;
                chk_yyyymm.Checked = false;
            }
        }

        protected void chk_yyyymm_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_yyyymm.Checked = true)
            {
                tb_yyyymm.Enabled = true;
                str_date.Enabled = false;
                chk_date.Checked = false;
            }
        }
    }
}