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
    public partial class AhnPopup : System.Web.UI.Page
    {
        SqlConnection _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["ekp"].ConnectionString);
        SqlCommand _sqlCmd;
        SqlDataReader _sqlDataReader;

        protected void Page_Load(object sender, EventArgs e)
        {
            ViewState["control_no"] = Request.QueryString["control_no"];
            string sqlQuery = string.Format(@"SELECT 
                                                replace(convert(nvarchar(max), content), '&quot;', '') as 제목
                                              FROM DAM010(NOLOCK) 
                                              WHERE 1=1
                                                AND control_no = '{0}'", ViewState["control_no"]);

            SqlStateCheck();
            _sqlCmd = new SqlCommand(sqlQuery, _sqlConn);
            _sqlDataReader = _sqlCmd.ExecuteReader();
            dataGrid.DataSource = _sqlDataReader;
            dataGrid.DataBind();
        }


        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
                _sqlDataReader.Close();
        }
    }
}