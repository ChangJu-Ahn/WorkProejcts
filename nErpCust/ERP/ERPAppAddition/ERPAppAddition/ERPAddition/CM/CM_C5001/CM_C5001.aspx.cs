using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using SRL.UserControls;
using System.Drawing;


namespace ERPAppAddition.ERPAddition.CM.CM_C5001
{
    public partial class CM_C5001 : System.Web.UI.Page
    {
        #region Global Variable Declaration (Sql Connection, Command, Reader)
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sqlCmd;
        SqlDataReader sqlDataReader;
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    InitDropDownPlant();
                    WebSiteCount();
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                    //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
                }
                finally
                {
                    if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                        sqlConn.Close();
                }
            }

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sqlConn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            try
            {
                SqlState_Check();
                sqlDataReader = GetValueProcedure();
                
                dt.Load(sqlDataReader);
                
                dgList.DataSource = dt;

                if (dt.Rows.Count > 0)
                    dgList.DataBind();
                else
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "OutputAlert('검색된 정보가 없습니다.');", true);

                //ReportViewerSetting(dt);
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
            }
            finally
            {
                if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
            }
        }

        #region ReportViewer Setting
        //protected void ReportViewerSetting(DataTable dt)
        //{
        //    ReportDataSource rds = new ReportDataSource("rdsContrastBom", dt);

        //    ReportViewer1.Reset();
        //    ReportViewer1.LocalReport.ReportPath = Server.MapPath("CM_C5001.rdlc");
        //    ReportViewer1.LocalReport.DisplayName = "투입현황대비_실사용량 [" + DateTime.Now.ToShortDateString() + "]";
        //    ReportViewer1.LocalReport.DataSources.Add(rds);
        //}
        #endregion

        #region Page Controls Setting (DropDownList)
        protected void InitDropDownPlant()
        {
            string sqlQuery = @"Select Top 101 PLANT_CD,PLANT_NM From   B_PLANT Where  PLANT_CD>= '' order by PLANT_CD";
            txtdate.Text = DateTime.Now.ToString("yyyyMM");

            SqlState_Check();
            sqlCmd = new SqlCommand(sqlQuery, sqlConn);
            sqlDataReader = sqlCmd.ExecuteReader();

            ddl_Plant.DataSource = sqlDataReader;
            ddl_Plant.DataValueField = "PLANT_CD";
            ddl_Plant.DataTextField = "PLANT_NM";
            ddl_Plant.DataBind();
        }
        #endregion

        protected SqlDataReader GetValueProcedure()
        {
            string strDate = txtdate.Text.ToString();
            string plant_Cd = ddl_Plant.SelectedValue.ToString().ToUpper();

            sqlCmd = new SqlCommand("dbo.CM_C5001_LIST", sqlConn);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandTimeout = 4000; 

            SqlParameter param1, param2;
            param1 = new SqlParameter("@YYYYMM", SqlDbType.VarChar, 6);
            param2 = new SqlParameter("@PLANT_CD", SqlDbType.VarChar, 4);

            param1.Value = strDate;
            param2.Value = plant_Cd;

            sqlCmd.Parameters.Add(param1);
            sqlCmd.Parameters.Add(param2);

            return sqlCmd.ExecuteReader();  

        }

        protected void SqlState_Check()
        {
            if (sqlConn == null || sqlConn.State == ConnectionState.Closed)
                sqlConn.Open();

            if (sqlDataReader != null && sqlDataReader.FieldCount > 0)
                sqlDataReader.Close();
        }

    }
}