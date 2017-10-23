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
using ERPAppAddition.ERPAddition.PM.PM_P1020;


namespace ERPAppAddition.ERPAddition.PM.PM_P1020
{
    public partial class PM_P1020 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ReportViewer1.Reset();
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

        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    ds.Tables[0].Load(dr);
                    dr.Close();
                    _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                    _reportViewer.LocalReport.DisplayName = "REPORT_" + dl_plant_cd.Text.Trim() + "BOM_" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = _ReportDataSourceName;
                    rds.Value = ds.Tables[0];
                    _reportViewer.LocalReport.DataSources.Add(rds);

                    _reportViewer.LocalReport.Refresh();
                }
                else
                    lblCnt.Text = "조회된 데이터가 없습니다.";

            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        // 부모창 - 조회버튼
        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            lblCnt.Text = "";
            string sql = "exec usp_p_put_child_item '" + dl_plant_cd.Text.Trim() + "'";
            DataSet_pm_p1020 dt1 = new DataSet_pm_p1020();
            ReportViewer1.Reset();
            ReportCreator(dt1, sql, ReportViewer1, "pm_p1020.rdlc", "DataSet1");
        }
    }

}