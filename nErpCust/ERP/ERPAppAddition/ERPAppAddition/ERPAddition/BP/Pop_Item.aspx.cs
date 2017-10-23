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
    public partial class Pop_Item : System.Web.UI.Page
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


            lblListCnt.Text = "0";

            string columns = Request.QueryString["columns"].ToString();
            string sql = Request.QueryString["sql"].ToString();
            if (Request.QueryString["itemcd"] != null && Request.QueryString["itemcd"].ToString() != "")
            {
                string item = Request.QueryString["itemcd"].ToString();
                txtItem_cd.Text = item;
            }
            
            //string search = Request.QueryString["search"].ToString();
          
            //string[] searchbar = search.Split(',');

            //lblitem_cd.Text = searchbar[0].Split(';')[1];
            //txtItem_cd.Text = searchbar[0].Split(';')[0];

            //lblItem_nm.Text = searchbar[1].Split(';')[1];
            //txtItem_nm.Text = searchbar[1].Split(';')[0];



            string[] colList = columns.Split(',');

            lblitem_cd.Text = colList[0];
            lblItem_nm.Text = colList[1];


            ViewState["columns"] = columns;
            ViewState["sql"] = sql;
        }

        protected void GetPartnerList()
        {
            //string sqlQuery = @"SELECT BP_CD AS '거래처코드', BP_NM AS '거래처이름', BP_RGST_NO AS '사업자등록번호' FROM B_BIZ_PARTNER WHERE USAGE_FLAG = 'Y'";
            string sqlQuery = ViewState["sql"].ToString();
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
                string[] columns = ViewState["columns"].ToString().Split(',');

                dgList.Columns.Clear();

                for (int i = 0; i < columns.Length; i++)
                {
                    TemplateField tf2 = new TemplateField();
                    tf2.HeaderText = columns[i];
                    tf2.HeaderStyle.BackColor = Color.FromArgb(0, 51, 153);
                    tf2.HeaderStyle.ForeColor = Color.White;
                    tf2.HeaderStyle.Width = Unit.Percentage(100 / dt.Rows.Count);
                    tf2.HeaderStyle.HorizontalAlign = HorizontalAlign.Left;
                    tf2.HeaderStyle.VerticalAlign = VerticalAlign.Middle;
                    tf2.ItemStyle.HorizontalAlign = HorizontalAlign.Left;
                    tf2.ItemStyle.VerticalAlign = VerticalAlign.Middle;
                    tf2.ItemStyle.Wrap = false;

                    dgList.Columns.Add(tf2);

                }
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

            string ITEM_CD = txtItem_cd.Text.ToString();
            string ITEM_NM = txtItem_nm.Text.ToString();
            string PartnerType = string.Empty;

           
            sb.AppendLine(ViewState["sql"].ToString());

            sb.AppendLine("WHERE 1=1");

            if (ITEM_CD.Length >= 1)
                //sb.AppendLine(string.Format("AND ITEM_CD >= '{0}'", ITEM_CD));
                sb.AppendLine(string.Format("AND ITEM_CD LIKE '%{0}%'", ITEM_CD));

            if (ITEM_NM.Length >= 1)
                sb.AppendLine(string.Format("AND ITEM_NM LIKE '%{0}%'", ITEM_NM));


            sb.AppendLine("ORDER BY 1");
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
                        e.Row.Cells[i].Text = string.Format("<a href =\"#\" onclick=\"PopDateDeliver('{0}', '{1}')\"> {0} </a>", rowView[0].ToString(), ConvertSpecialStr(rowView[1].ToString()));
                    else
                        e.Row.Cells[i].Text = rowView[i].ToString();
                }
            }
        }

        private string ConvertSpecialStr(string str)
        {
             str = str.Replace("<","&lt;");
             str = str.Replace(">","&gt;");
             str = str.Replace("\"","&quot;");
             str = str.Replace("\'","&#39;");
             str = str.Replace("\\n","<br />");


            return str;
        }

    }
}