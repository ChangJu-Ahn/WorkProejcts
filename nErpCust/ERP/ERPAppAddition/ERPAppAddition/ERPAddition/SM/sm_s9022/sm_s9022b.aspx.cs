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

namespace ERPAppAddition.ERPAddition.SM.sm_s9022
{
    public partial class sm_s9022b : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        //string sql_cust_cd;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);        
        SqlCommand sql_cmd = new SqlCommand();
        DataSet ds = new DataSet();

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleDataReader dr;
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
                sbSQL.Append("   and PLAN_YEAR >= '2015'                              \n");
                sbSQL.Append("   and PLAN_YEAR <= to_char(sysdate, 'yyyy')            \n");
                sbSQL.Append("   and PLAN_YEAR || LPAD(PLAN_MONTH, 2, 0) <= to_char(sysdate-1, 'YYYYMM') \n");
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
            if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"조회년을 선택해주세요.\");";
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
                sql_cmd.CommandTimeout = 0;

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);

                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                    //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
                }
                sql_conn.Close();

                /*seq 가 a는 필수 항목 없는경우 조회 불가*/
                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s9022b.rdlc");
                ReportViewer1.LocalReport.DisplayName = "재고출고관리 조회" + DateTime.Now.ToShortDateString();

                ReportDataSource rds = new ReportDataSource();
                DataTable dt1 = ds.Tables["DataSet1"].Copy();
                //dt1.DefaultView.RowFilter = "SQ LIKE '1%'";
                rds.Name = "DataSet1";
                rds.Value = dt1;
                ReportViewer1.LocalReport.DataSources.Add(rds);



                ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
            }
        }

        private string getSQL()
        {
            string strDATE = tb_yyyymm.Text;

            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("USP_S_BOM_CHK03  '" + strDATE + "' \n");
            //sbSQL.Append("select 'aaa' as item_cd  \n");
            return sbSQL.ToString();
        }
    }
}