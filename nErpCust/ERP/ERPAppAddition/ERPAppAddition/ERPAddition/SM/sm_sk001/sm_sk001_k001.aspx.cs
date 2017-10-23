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
namespace ERPAppAddition.ERPAddition.SM.sm_sk001
{
    public partial class sm_sk001_k001 : System.Web.UI.Page
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
        DataTable dtYYYYMM = new DataTable();

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

        private void setMonth()
        {
            conn.Open();
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select PLAN_YEAR || lpad(PLAN_MONTH, 2, 0) AS M_MONTH   \n");
                sbSQL.Append("  from CALENDAR                                         \n");
                sbSQL.Append(" where PLANT = 'CCUBEDIGITAL'                           \n");
                sbSQL.Append("   and PLAN_YEAR >= '2016'                              \n");
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

                ///*DDL 프로젝트 조회조건*/
                //sql_conn.Open();
                //sql_cmd = sql_conn.CreateCommand();
                //sql_cmd.CommandType = CommandType.Text;
                //sql_cmd.CommandText = "SELECT DISTINCT MINOR_CD, MINOR_NM from B_MINOR where MAJOR_CD = 'A9003' and  MINOR_NM like '%소재]%' or MINOR_NM like '기술원%' or MINOR_NM like '%부품]%' ORDER BY 2 ";
                //SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                //da.Fill(ds, "DataSet1");

                //if (ds.Tables["DataSet1"].Rows.Count > 0)
                //{
                //    DDL_PROJ.DataSource = ds.Tables["DataSet1"];
                //    DDL_PROJ.DataValueField = "MINOR_CD";
                //    DDL_PROJ.DataTextField = "MINOR_NM";
                //    DDL_PROJ.DataBind();                    
                //}
                //sql_conn.Close();                


                /*DDL 연구소 조회조건*/
                sql_conn.Open();
                sql_cmd2 = sql_conn.CreateCommand();
                sql_cmd2.CommandType = CommandType.Text;
                sql_cmd2.CommandText = "SELECT DISTINCT GRP_CD1, GRP_NM1 + '(' +  GRP_CD1 + ')'  as GRP_NM1 fROM T_TEC_IN_USER  where convert(varchar, ver_no) + convert(varchar, yyyymm) in(select convert(varchar, ver_no) + convert(varchar, yyyymm) from ( select max(VER_NO) as ver_no, YYYYMM from T_TEC_IN_USER group by YYYYMM)a ) ORDER BY 2";
                SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                da2.Fill(ds, "DataSet2");

                if (ds.Tables["DataSet2"].Rows.Count > 0)
                {
                    DDL_DEPT.DataSource = ds.Tables["DataSet2"];
                    DDL_DEPT.DataValueField = "GRP_CD1";
                    DDL_DEPT.DataTextField = "GRP_NM1";
                    DDL_DEPT.DataBind();                    
                }
                sql_conn.Close();

                DDL_DEPT_CHANGE();

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
            if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"조회년을 선택해주세요.\");";
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

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sk001_k001.rdlc");
                    ReportViewer1.LocalReport.DisplayName = tb_yyyymm.Text + "_개발비용명세서조회(기술원)" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = ds.Tables["DataSet1"];
                    ReportViewer1.LocalReport.DataSources.Add(rds);                   


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
            string dept = DDL_DEPT.SelectedValue;
            string proj = DDL_PROJ.SelectedValue;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            if (CheckBox1.Checked)
            {
                sbSQL.Append(" EXEC	USP_AGL_KISULL01  '" + date + "', '" + dept + "','" + proj + "' \n");
            }
            else
            { //프로젝트 코드를 넣어서 조회를 안할때 체크 푼다. 
                sbSQL.Append(" EXEC	USP_AGL_KISULL02  '" + date + "', '" + dept + "' \n");
            }
            
            return sbSQL.ToString();
        }

        protected void DDL_DEPT_SelectedIndexChanged(object sender, EventArgs e)
        {
            DDL_DEPT_CHANGE();
        }

        private void DDL_DEPT_CHANGE()
        {
            string str = DDL_DEPT.SelectedItem.Text;

            if (str.Length > 11)
            {
                if (str.IndexOf("소재") > 0) str = "%소재%";
                if (str.IndexOf("부품") > 0) str = "%부품%";
            }
            else
            {
                str = "기술원(공통)";
            }

            /*DDL 프로젝트 조회조건*/
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = "SELECT DISTINCT MINOR_CD, MINOR_NM from B_MINOR where MAJOR_CD = 'A9003' and  MINOR_NM like '" + str + "' ORDER BY 2";
            SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
            da.Fill(ds, "DataSet1");

            if (ds.Tables["DataSet1"].Rows.Count > 0)
            {
                DDL_PROJ.DataSource = ds.Tables["DataSet1"];
                DDL_PROJ.DataValueField = "MINOR_CD";
                DDL_PROJ.DataTextField = "MINOR_NM";
                DDL_PROJ.DataBind();
            }
            sql_conn.Close();
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(CheckBox1.Checked){
                DDL_PROJ.Enabled = true;
            }else{
                DDL_PROJ.Enabled = false;
            }
            
        }
    }
}