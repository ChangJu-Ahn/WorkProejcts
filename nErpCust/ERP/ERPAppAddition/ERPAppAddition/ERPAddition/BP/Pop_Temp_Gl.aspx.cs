using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.BP
{
    public partial class Pop_Temp_Gl : System.Web.UI.Page
    {
        SqlConnection _sqlConn;
        //SqlDataReader _sqlDataReader;
        DataTable dtReturn;
        SqlDataAdapter Sqladapter;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                DateTime setDate = DateTime.Today.AddDays(-7);
                tb_fr_yyyymmdd.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00") + setDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
            }
            if (Request.QueryString["dbName"] != null && Request.QueryString["dbName"].ToString() != "")
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings[Request.QueryString["dbName"].ToString()].ConnectionString);
            else
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

            if(Request.QueryString["AcctNo"] != null && Request.QueryString["AcctNo"].ToString() != "")
            {
                ViewState["AcctNo"] = Request.QueryString["AcctNo"].ToString();
            }
            else
            {
                ViewState["AcctNo"] = "";
            }

            dtReturn = new DataTable();
        }

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            Sreach();
        }

        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            //if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
            //    _sqlDataReader.Close();
        }

        private void Sreach()
        {

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("");
            sb.AppendLine(" SELECT DISTINCT A.TEMP_GL_NO ");
            sb.AppendLine(" 	,B.DEPT_NM ");
            sb.AppendLine(" 	,A.TEMP_GL_DT ");
            sb.AppendLine(" 	,A.REF_NO ");
            sb.AppendLine(" 	,D.USR_NM ");
            sb.AppendLine(" 	,A.TEMP_GL_DESC ");
            sb.AppendLine(" 	,CASE WHEN A.DR_LOC_AMT = 0 ");
            sb.AppendLine(" 			THEN A.CR_LOC_AMT ");
            sb.AppendLine(" 	      ELSE A.DR_LOC_AMT ");
            sb.AppendLine(" 	  END AS AMT ");
            sb.AppendLine(" FROM A_TEMP_GL A WITH(NOLOCK) ");
            sb.AppendLine("    LEFT OUTER JOIN B_ACCT_DEPT B WITH(NOLOCK) ");
            sb.AppendLine("    ON A.DEPT_CD = B.DEPT_CD ");
            sb.AppendLine("    AND A.ORG_CHANGE_ID = B.ORG_CHANGE_ID ");
            sb.AppendLine("    LEFT OUTER JOIN Z_USR_MAST_REC D WITH(NOLOCK) ");
            sb.AppendLine("    ON A.INSRT_USER_ID = D.USR_ID ");
            sb.AppendLine("    LEFT OUTER JOIN A_TEMP_GL_ITEM I WITH(NOLOCK) ");
            sb.AppendLine("    ON A.TEMP_GL_NO = I.TEMP_GL_NO ");
            sb.AppendLine(" WHERE 1=1 ");
            
            sb.AppendLine(" 	AND A.TEMP_GL_DT >= '"+tb_fr_yyyymmdd.Text+"' ");
            sb.AppendLine(" 	AND A.TEMP_GL_DT <= '" + tb_to_yyyymmdd.Text + "' ");
            sb.AppendLine(" 	AND A.GL_INPUT_TYPE <> 'TD' ");

            if(txtTempGL_NO.Text.Length > 0)
            {
                sb.AppendLine(" 	AND A.TEMP_GL_NO = '" + txtTempGL_NO.Text + "' ");
            }

            if (rdoConf_Y.Checked)
            {
                sb.AppendLine(" 	AND A.CONF_FG = 'C' ");
            }

            if(txtBizArea_CD.Text.Length > 0)
            {
                sb.AppendLine(" 	AND A.biz_area_cd = '" + txtBizArea_CD.Text + "' ");
            }

            if(txtDept_cd.Text.Length > 0)
            {
                sb.AppendLine(" 	AND A.dept_cd = '" + txtDept_cd.Text + "' ");
            }

            if (txtTempGl_Desc.Text.Length > 0)
            {
                sb.AppendLine(" 	AND A.TEMP_GL_DESC LIKE '%" + txtTempGl_Desc.Text + "%' ");
            }
            if (ViewState["AcctNo"].ToString().Length > 0)
            {
                sb.AppendLine(" 	AND I.ACCT_CD LIKE '" + ViewState["AcctNo"].ToString() + "%' ");
            }

            sb.AppendLine(" ORDER BY A.TEMP_GL_NO ASC ");
            sb.AppendLine(" 	,B.DEPT_NM ASC ");
            sb.AppendLine(" 	,A.TEMP_GL_DT ASC ");
            sb.AppendLine(" 	,A.REF_NO ASC ");
            sb.AppendLine(" 	,D.USR_NM ASC ");
            sb.AppendLine(" 	,A.TEMP_GL_DESC ASC ");

            try
            {
                SqlStateCheck();

                dtReturn = new DataTable();
                Sqladapter = new SqlDataAdapter(sb.ToString(), _sqlConn);
                Sqladapter.Fill(dtReturn);

                SetDataBind(dtReturn);
            }
            catch (Exception ex)
            {
                //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //UpdatePanel로 인해 MessageBox 출력 안됨, 화면단 메소드를 호출할 수 있도록 변경
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
            }
            finally
            {
                if (_sqlConn != null || _sqlConn.State == ConnectionState.Open)
                    _sqlConn.Close();
            }
        }

        protected void SetDataBind(DataTable dt)
        {
            if (dt.Rows.Count > 0)
            {
                lblListCnt.Text = dt.Rows.Count.ToString();
                dgList.DataSource = dt;
                dgList.DataBind();
            }
            else
            {
                lblListCnt.Text = "0";
                dgList.Controls.Clear();
            }
        }


        protected void dgList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            dgList.PageIndex = e.NewPageIndex;
            btnSelect_Click(sender, e);
        }

        protected void dgList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataRowView rowView = (DataRowView)e.Row.DataItem;
            string[] tempArray;

            if (rowView != null)
            {
                tempArray = new string[rowView.Row.ItemArray.Length];

                for (int i = 0; i < tempArray.Length; i++)
                {
                    if (i == 0)
                        e.Row.Cells[i].Text = string.Format("<a href =\"#\" onclick=\"PopDateDeliver('{0}', '{1}')\"> {0} </a>", rowView[0].ToString(), rowView[1].ToString());
                    else
                        e.Row.Cells[i].Text = rowView[i].ToString();
                }
            }
        }
    }
}