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
using ERPAppAddition.ERPAddition.AM.AM_A4002;

namespace ERPAppAddition.ERPAddition.AM.AM_A4002
{
    public partial class AM_A4002 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_led"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ReportViewer1.Reset();
                setDDL();
                WebSiteCount();
            }

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void setDDL()
        {
            DataTable BankDt = DDL(getDDLSQL());
            if (BankDt.Rows.Count > 0)
            {
                ddlDATE.DataTextField = "ALLC_DT";
                ddlDATE.DataValueField = "ALLC_DT";
                ddlDATE.DataSource = BankDt;
                ddlDATE.DataBind();
                ddlDATE.SelectedIndex = 0;
            }
        }

        private string getDDLSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT REPLACE(CONVERT(VARCHAR, L.ALLC_DT, 111), '/', '-') AS ALLC_DT       		\n");
            sbSQL.Append("FROM A_CLS_ACCT_ITEM A                                                              \n");
            sbSQL.Append("INNER JOIN A_GL_ITEM B ON A.GL_NO=B.GL_NO AND A.GL_SEQ=B.ITEM_SEQ                   \n");
            sbSQL.Append("INNER JOIN B_MINOR C ON C.MAJOR_CD='A1050' AND C.MINOR_CD = 'U9'                    \n");
            sbSQL.Append("LEFT JOIN AV_A_OPEN_ACCT_BP_CD D ON A.GL_NO=D.GL_NO AND A.GL_SEQ=D.GL_SEQ           \n");
            sbSQL.Append("LEFT JOIN A_ACCT E ON B.ACCT_CD=E.ACCT_CD                                           \n");
            sbSQL.Append("INNER JOIN B_MINOR G ON G.MAJOR_CD = 'A1012' AND G.MINOR_CD=E.BAL_FG                \n");
            sbSQL.Append("LEFT JOIN B_ACCT_DEPT H ON B.ORG_CHANGE_ID=H.ORG_CHANGE_ID AND B.DEPT_CD=H.DEPT_CD  \n");
            sbSQL.Append("LEFT JOIN B_BIZ_AREA J ON B.BIZ_AREA_CD=J.BIZ_AREA_CD                               \n");
            sbSQL.Append("INNER JOIN A_OPEN_ACCT K ON A.GL_NO=K.GL_NO AND A.GL_SEQ=K.GL_SEQ                   \n");
            sbSQL.Append("inner join A_ALLC_HDR L ON A.CLS_NO=L.ALLC_NO                                       \n");
            sbSQL.Append("left join HAA010T M on M.EMP_NO= K.MGNT_VAL1                                        \n");
            sbSQL.Append("where a.acct_cd = '21100903'   and K.MGNT_TYPE = '9'                                \n");
            sbSQL.Append("AND L.ALLC_DT > CONVERT(DATE, GETDATE()-365)                                        \n");
            sbSQL.Append("GROUP BY  L.ALLC_DT                                                                 \n");
            sbSQL.Append("order by L.ALLC_DT desc                                                             \n");
            return sbSQL.ToString();
        }

        private DataTable DDL(string SQL)
        {
            DataTable resultDt = new DataTable();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = SQL;

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(resultDt);
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

            return resultDt;
        }


        protected void Load_btn_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "usp_bankmasterlist_view";
            cmd.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@gl_dt", SqlDbType.VarChar, 10);


            param1.Value = ddlDATE.Text;

            cmd.Parameters.Add(param1);


            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rp_am_a4002.rdlc");
                ReportViewer1.LocalReport.DisplayName = "은행이체리스트(LED 직원)" + ddlDATE.Text + DateTime.Now.ToShortDateString();

                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportViewer1.LocalReport.Refresh();

                //UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
    }
}


