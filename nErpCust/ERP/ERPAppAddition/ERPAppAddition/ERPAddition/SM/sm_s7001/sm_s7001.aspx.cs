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



namespace ERPAppAddition.ERPAddition.SM.sm_s7001
{
    public partial class sm_s7001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        
        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleConnection conn_if = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_UNIERP"].ConnectionString);

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        SqlCommand sql_cmd2 = new SqlCommand();
        SqlCommand sql_cmd3 = new SqlCommand();
        SqlCommand sql_cmd4 = new SqlCommand();        

        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr, ndr;        
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();        
        DataTable dtYYYYMM= new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
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
                sbSQL.Append("   and PLAN_YEAR || LPAD(PLAN_MONTH, 2, 0) <= to_char(sysdate, 'YYYYMM') \n");
                sbSQL.Append(" group by PLAN_YEAR, PLAN_MONTH                         \n");
                sbSQL.Append(" order by 1 desc                                             \n");
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
            if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null){
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

                    sql_cmd2 = sql_conn.CreateCommand();
                    sql_cmd2.CommandType = CommandType.Text;
                    sql_cmd2.CommandText = getSQL2();

                    sql_cmd3 = sql_conn.CreateCommand();
                    sql_cmd3.CommandType = CommandType.Text;
                    sql_cmd3.CommandText = getSQL3();

                    sql_cmd4 = sql_conn.CreateCommand();
                    sql_cmd4.CommandType = CommandType.Text;
                    sql_cmd4.CommandText = getSQL4();

                    DataTable dt = new DataTable();                    
                    try
                    {
                        SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                        da.Fill(ds, "DataSet1");

                        SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                        da2.Fill(ds, "DataSet2");

                        SqlDataAdapter da3 = new SqlDataAdapter(sql_cmd3);
                        da3.Fill(ds, "DataSet3");

                        SqlDataAdapter da4 = new SqlDataAdapter(sql_cmd4);
                        da4.Fill(ds, "DataSet4");
                    }
                    catch (Exception ex)
                    {
                        if (sql_conn.State == ConnectionState.Open)
                            sql_conn.Close();
                    }
                    sql_conn.Close();  
                    
                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s7001.rdlc");
                    ReportViewer1.LocalReport.DisplayName = tb_yyyymm.Text + "_EM 매출현황" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = ds.Tables["DataSet1"];
                    ReportViewer1.LocalReport.DataSources.Add(rds);
                                        
                    /*원그래프*/
                    ReportDataSource rds2 = new ReportDataSource();
                    rds2.Name = "DataSet2_1";
                    DataTable dt2_1 = ds.Tables["DataSet2"].Copy();
                    dt2_1.DefaultView.RowFilter = "RW = '1'";
                    rds2.Value = dt2_1.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(rds2);

                    ReportDataSource rds2_2 = new ReportDataSource();
                    rds2_2.Name = "DataSet2_2";
                    DataTable dt2_2 = ds.Tables["DataSet2"].Copy();
                    dt2_2.DefaultView.RowFilter = "RW = '2'";
                    rds2_2.Value = dt2_2.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(rds2_2);

                    ReportDataSource rds2_3 = new ReportDataSource();
                    rds2_3.Name = "DataSet2_3";
                    DataTable dt2_3 = ds.Tables["DataSet2"].Copy();
                    dt2_3.DefaultView.RowFilter = "RW = '3'";
                    rds2_3.Value = dt2_3.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(rds2_3);

                    ReportDataSource rds3 = new ReportDataSource();
                    rds3.Name = "DataSet3";
                    rds3.Value = ds.Tables["DataSet3"];
                    ReportViewer1.LocalReport.DataSources.Add(rds3);

                    ReportDataSource rds4 = new ReportDataSource();
                    rds4.Name = "DataSet4";
                    rds4.Value = ds.Tables["DataSet4"];
                    ReportViewer1.LocalReport.DataSources.Add(rds4);   

                
                    ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x                    

                    DataTable dt2 = new DataTable();

                    dt2.Load(ndr);
                    if (dt2.Rows.Count > 0)
                    {
                        ReportViewer1.LocalReport.Refresh();
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
            sbSQL.Append("	USP_S_BP_MONTH  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL2()
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MONTH_GRAPH  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL3()
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MONTH_DET  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL4()
        {
            string strValue = tb_yyyymm.Items[tb_yyyymm.SelectedIndex].Value;
            string date = tb_yyyymm.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MONTH_CHART  '" + date + "'\n");
            return sbSQL.ToString();
        }
    }
}