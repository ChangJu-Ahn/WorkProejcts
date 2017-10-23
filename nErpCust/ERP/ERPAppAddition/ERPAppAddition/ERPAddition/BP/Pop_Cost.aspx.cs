using System;
using System.Data;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using Microsoft.Reporting.WebForms;

namespace ERPAppAddition.ERPAddition.BP
{
    public partial class Pop_Cost : System.Web.UI.Page
    {
        SqlConnection _sqlConn;
        //SqlDataReader _sqlDataReader;
        DataTable dtReturn;
        SqlDataAdapter Sqladapter;

        protected void Page_Init(object sender, EventArgs e)
        {
            //InitParam();
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                //GetPartnerList(); //초기에 모든 데이터 불러옴
                InitParam();
                dtReturn = new DataTable();
                Sreach();
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
            }
            finally
            {
                if (_sqlConn != null || _sqlConn.State == ConnectionState.Open)
                    _sqlConn.Close();
            }
        }

        protected void InitParam()
        {
            if (Request.QueryString["dbName"] != null && Request.QueryString["dbName"].ToString() != "")
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings[Request.QueryString["dbName"].ToString()].ConnectionString);
            else
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

            if (Request.QueryString["userId"] != null && Request.QueryString["userId"].ToString() != "")
                ViewState["userId"] = Request.QueryString["userId"].ToString();
            else
                ViewState["userId"] = "DEV";

            lblListCnt.Text = "0";

            if (Request.QueryString["userId"] != null && Request.QueryString["userId"].ToString() != "")
            {
                string search = Request.QueryString["search"].ToString();

                 txtPartnerCD.Text = search.Split(',')[0];
                 txtPartnerNm.Text = search.Split(',')[1];
            }


            

        }

        protected void GetPartnerList()
        {
            string sqlQuery = @"SELECT BP_CD AS '거래처코드', BP_NM AS '거래처이름', BP_RGST_NO AS '사업자등록번호' FROM B_BIZ_PARTNER WHERE USAGE_FLAG = 'Y'";
            dtReturn = new DataTable();

            SqlStateCheck();
            Sqladapter = new SqlDataAdapter(sqlQuery, _sqlConn);
            Sqladapter.Fill(dtReturn);

            SetDataBind(dtReturn);
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

        #region DB Connection and DataReader State Check
        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            //if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
            //    _sqlDataReader.Close();
        }
        #endregion

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            Sreach();
        }


        private void Sreach()
        {
            string sqlQuery = GetSelectQuery();


            try
            {
                SqlStateCheck();

                dtReturn.Clear(); //dt를 재사용 하기 위한 초기화 실시
                Sqladapter = new SqlDataAdapter(sqlQuery, _sqlConn);
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
        protected string GetSelectQuery()
        {
            StringBuilder sb = new StringBuilder();

            string PartnerCD = txtPartnerCD.Text.ToString();
            string PartnerNm = txtPartnerNm.Text.ToString();
            string PartnerType = string.Empty;

            if (rdoCS.Checked == true)
                PartnerType = "CS";
            else if (rdoS.Checked == true)
                PartnerType = "S";
            else if (rdoC.Checked == true)
                PartnerType = "C";
            else
                PartnerType = "ALL";

            sb.AppendLine("SELECT BP_CD AS '거래처코드', BP_NM AS '거래처이름', BP_RGST_NO AS '사업자등록번호' FROM B_BIZ_PARTNER WHERE USAGE_FLAG = 'Y'");

            if (PartnerCD.Length >= 1)
                //sb.AppendLine(string.Format("AND BP_CD >= '{0}'", PartnerCD));
                sb.AppendLine(string.Format("AND BP_CD LIKE '%{0}%'", PartnerCD));

            if (PartnerNm.Length >= 1)
                //sb.AppendLine(string.Format("AND BP_NM >= '{0}'", PartnerNm));
                sb.AppendLine(string.Format("AND BP_NM LIKE '%{0}%'", PartnerNm));

            if (PartnerType != "ALL")
                sb.AppendLine(string.Format("AND BP_TYPE = '{0}'", PartnerType));

            return sb.ToString();
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
                    if (i==0)
                        e.Row.Cells[i].Text = string.Format("<a href =\"#\" onclick=\"PopDateDeliver('{0}', '{1}')\"> {0} </a>", rowView[0].ToString(), rowView[1].ToString());
                    else
                        e.Row.Cells[i].Text = rowView[i].ToString();
                }
            }
        }

    }
}