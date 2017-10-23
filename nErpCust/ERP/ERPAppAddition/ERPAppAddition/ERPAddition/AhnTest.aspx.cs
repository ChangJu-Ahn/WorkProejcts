using System.Data.Common;
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
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Threading;
using System.Drawing;


namespace ERPAppAddition.ERPAddition
{
    public partial class AhnTest : System.Web.UI.Page
    {
        SqlConnection _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["ekp"].ConnectionString);
        SqlCommand _sqlCmd;
        SqlDataReader _sqlDataReader;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                InitDropDownPlant();
            }

        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            DataTable dtReturn = new DataTable();
            SqlDataAdapter Sqladapter;

            string contents = txtContent.Text;
            string erpFlag = ddlList.SelectedValue;
            string sqlQuery = string.Format(@"SELECT 
                                                control_no as 문서키
                                                , doc_no as 문서번호 
                                                , subject as 제목
                                                , write_dt as 작성일
                                              FROM DAM010(NOLOCK) 
                                              WHERE 1=1
                                                AND ERP_FLAG = '{0}'
                                                AND SUBJECT LIKE '%{1}%'
                                               ORDER BY CONTROL_NO", erpFlag, contents);

            SqlStateCheck();
            Sqladapter = new SqlDataAdapter(sqlQuery, _sqlConn);
            Sqladapter.Fill(dtReturn);
            gridView.DataSource = dtReturn;
            gridView.DataBind();
        }

        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
                _sqlDataReader.Close();
        }

        protected void InitDropDownPlant()
        {
            string sqlQuery = @"
                SELECT 
	                CASE WHEN UPPER(ERP_FLAG) = '' THEN '보고서'
	                ELSE UPPER(ERP_FLAG) 
	                END AS FLAG
                    ,ERP_FLAG
                FROM DAM010 (NOLOCK)
                GROUP BY ERP_FLAG
            ";

            SqlStateCheck();
            _sqlCmd = new SqlCommand(sqlQuery, _sqlConn);
            _sqlDataReader = _sqlCmd.ExecuteReader();

            ddlList.DataSource = _sqlDataReader;
            ddlList.DataValueField = "ERP_FLAG";
            ddlList.DataTextField = "FLAG";
            ddlList.DataBind();
        }

        protected void gridView_ItemDataBound(object sender, GridViewRowEventArgs e)
        {
            DataRowView dv = (DataRowView)e.Row.DataItem;
            if (dv == null) return;

            e.Row.Cells[0].Text = dv["문서키"].ToString();
            e.Row.Cells[1].Text = dv["문서번호"].ToString();
            e.Row.Cells[2].Text = "<a href='AhnPopup.aspx?CONTROL_NO=" + dv["문서키"].ToString() + "' target='_black' >" + dv["제목"].ToString() + "</a>";

            //ekp화면에 링크를 걸어봤으나, 로그인이 필요하고 또한 첨부했던 파일이 날라가버려서 다운로드도 불가하다고 함, 그러기에 그냥 기존과 동일하게 변경
            //e.Row.Cells[2].Text = "<a href=\"javascript:on_view1(\'" + dv["문서키"].ToString() + "\', \'Y\')\" shape= \"\">" + dv["제목"].ToString() + "</a>";

            e.Row.Cells[3].Text = dv["작성일"].ToString();
        }

        protected void gridView_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridView.PageIndex = e.NewPageIndex;
            btnSearch_Click(sender, e);
        }


    }
}