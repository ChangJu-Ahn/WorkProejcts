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
using ERPAppAddition.ERPAddition.SM;

namespace ERPAppAddition.ERPAddition.SM.sm_s4001
{
    public partial class sm_s4001_select : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string sql;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                ReportViewer1.Reset();
            }
        }


        protected void btn_select_Click(object sender, EventArgs e) //조회버튼 클릭
        {
            ReportViewer1.Reset();
            sql = "SELECT a.cust_nm,a.item_nm,a.item_gp,a.size,a.process_type,a.route,a.pkg_type,a.remark,a.plan_mm,a.qty from S_FCST_QTY_IMPORT a";


            ds_sm_s4001_qty dt1 = new ds_sm_s4001_qty();
            ReportViewer1.Reset();
            ReportCreator(dt1, sql, ReportViewer1, "rp_sm_s4001_qty.rdlc", "DataSet1");


        }
        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer ReportViewer1, string _ReportName, string _ReportDataSourceName)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = _Query;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                ReportViewer1.LocalReport.ReportPath = Server.MapPath(_ReportName);

                ReportViewer1.LocalReport.DisplayName = "REPORT_";
                ReportDataSource rds = new ReportDataSource();
              
                rds.Name = "DataSet1";
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                ReportViewer1.LocalReport.DataSources.Add(rds);
               
                ReportViewer1.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
    }
}
   
