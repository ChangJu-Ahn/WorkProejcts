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

namespace ERPAppAddition.ERPAddition.MM.MM_MM001
{
    public partial class MM_MM001 : System.Web.UI.Page
    {

        #region Global Variable Declaration (Sql Connection, Command, Reader)
        SqlConnection _sqlConn;
        SqlCommand _sqlCmd;
        SqlDataReader _sqlDataReader;
        int cnt;

        #endregion

        protected void Page_Init(object sender, EventArgs e)
        {
            InitParam();
        }

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
                    //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //UpdatePanel로 인해 MessageBox 출력 안됨, 화면단 메소드를 호출할 수 있도록 변경
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                }
                finally
                {
                    if (_sqlConn != null || _sqlConn.State == ConnectionState.Open)
                        _sqlConn.Close();
                }
            }
        }

        #region DB Connection and DataReader State Check
        protected void SqlStateCheck()
        {
            if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
                _sqlConn.Open();

            if (_sqlDataReader != null && _sqlDataReader.FieldCount > 0)
                _sqlDataReader.Close();
        }
        #endregion

        #region Page Controls Setting (DropDownList, MultiCheckCombo)
        protected void InitDropDownPlant()
        {
            string sqlQuery = @"Select Top 101 PLANT_CD,PLANT_NM From   B_PLANT Where  PLANT_CD>= '' order by PLANT_CD";

            SqlStateCheck();
            _sqlCmd = new SqlCommand(sqlQuery, _sqlConn);
            _sqlDataReader = _sqlCmd.ExecuteReader();

            ddl_Plant.DataSource = _sqlDataReader;
            ddl_Plant.DataValueField = "PLANT_CD";
            ddl_Plant.DataTextField = "PLANT_NM";
            ddl_Plant.DataBind();
        }

        protected void InitParam()
        {
            if (Request.QueryString["dbName"] != null && Request.QueryString["dbName"].ToString() != "")
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings[Request.QueryString["db"].ToString()].ConnectionString);
            else
                _sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

            if (Request.QueryString["userId"] != null && Request.QueryString["userId"].ToString() != "")
                ViewState["userId"] = Request.QueryString["userId"].ToString();
            else
                ViewState["userId"] = "DEV";

            lblerpName.Text = _sqlConn.Database.ToString().ToUpper();
            lblListCnt.Text = "0";
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = _sqlConn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }
        #endregion

        #region MSSQL Query Setting
        protected string GetSelectQuery()
        {
            StringBuilder sb = new StringBuilder();

            string DocumentNo = txtDocument.Text.ToString();
            string ReqPrsn = txtReqPrsn.Text.ToString();
            string FromDay = txtdate_From.Text.ToString();
            string ToDay = txtdate_To.Text.ToString();
            string PlantCd = ddl_Plant.SelectedValue.ToString().ToUpper();
            string FlagType = ddlFlagType.SelectedValue.ToString().ToUpper();
            string MROType = ddlMROFlag.SelectedValue.ToString().ToUpper();

            sb.AppendLine(@"
                            SELECT
                                A.PR_GW_APP_NO
                              , CASE WHEN A.PR_NO = '' THEN '-'
	                              ELSE A.PR_NO
                                END AS PR_NO
                              , A.PLANT_CD
                              , A.ITEM_CD
                              , A.ITEM_NM
                              , A.REQ_QTY
                              , A.REQ_UNIT
                              , CONVERT(VARCHAR(15), A.REQ_DT, 23) as REQ_DT
                              , CONVERT(VARCHAR(15), A.DLVY_DT, 23) as DLVY_DT
                              , (SELECT USR_NM FROM Z_USR_MAST_REC WHERE USR_ID = A.REQ_PRSN) AS REQ_PRSN
                              , CASE WHEN A.ERP_APPLY_FG1 = 'E' THEN A.ERP_APPLY_FG1
                                     WHEN A.ERP_APPLY_FG2 = 'E' THEN A.ERP_APPLY_FG2
                                     WHEN A.ERP_APPLY_FG3 = 'E' THEN A.ERP_APPLY_FG3
                                     WHEN A.ERP_APPLY_FG4 = 'E' THEN A.ERP_APPLY_FG4
                                     WHEN A.ERP_APPLY_FG5 = 'E' THEN A.ERP_APPLY_FG5
                                ELSE 'Y'
                                END AS ERP_APPLY_FG
                              , A.ERP_APPLY_ERROR
                              FROM T_IF_RCV_PUR_REQ_KO441 A (NOLOCK)
                              WHERE 1=1
                             AND PR_GW_APP_NO <> ''
                           ");

            if (DocumentNo.Length > 0)
                sb.AppendLine(string.Format("AND A.PR_GW_APP_NO LIKE '%{0}%'", DocumentNo));

            if (ReqPrsn.Length > 0)
                sb.AppendLine(string.Format("AND A.REQ_PRSN  = (SELECT USR_ID FROM Z_USR_MAST_REC WHERE USR_NM LIKE '%{0}%')", ReqPrsn));

            if (PlantCd.Length > 0)
                sb.AppendLine(string.Format("AND A.PLANT_CD = '{0}'", PlantCd));

            if (FromDay.Length > 0)
                sb.AppendLine(string.Format("AND A.INSRT_DT >= '{0}'", FromDay));

            if (ToDay.Length > 0)
                sb.AppendLine(string.Format("AND A.INSRT_DT <= '{0}'", ToDay));

            if (MROType.Length > 0)
                sb.AppendLine(string.Format("AND A.MRO_ITEM_FG = '{0}'", MROType));
            else
                sb.AppendLine("AND A.MRO_ITEM_FG = 'N'");

            if (FlagType == "Y")
                sb.AppendLine("AND (ERP_APPLY_FG1 = 'Y' AND ERP_APPLY_FG2 = 'Y' AND ERP_APPLY_FG3 = 'Y' AND ERP_APPLY_FG4 = 'Y' AND ERP_APPLY_FG5 = 'Y')");
            else if (FlagType == "E")
                sb.AppendLine(string.Format("AND (ERP_APPLY_FG1 = '{0}' OR ERP_APPLY_FG2 = '{0}' OR ERP_APPLY_FG3 = '{0}' OR ERP_APPLY_FG4 = '{0}' OR ERP_APPLY_FG5 = '{0}')", FlagType));

            if (ddlMROFlag.SelectedValue != "Y")
            {
                sb.AppendLine(@"UNION ALL
                                SELECT
                                    A.GW_APP_NO
                                  , CASE WHEN PR_NO = '' THEN '-'
	                                  ELSE A.PR_NO
                                    END AS PR_NO
                                  , A.PLANT_CD
                                  , A.ITEM_CD
                                  , (SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = A.ITEM_CD) AS ITEM_NM
                                  , A.REQ_QTY
                                  , A.REQ_UNIT
                                  , CONVERT(VARCHAR(15), A.REQ_DT, 23) as REQ_DT
                                  , CONVERT(VARCHAR(15), A.DLVY_DT, 23) as DLVY_DT
                                  , (SELECT USR_NM FROM Z_USR_MAST_REC WHERE USR_ID = A.REQ_PRSN) AS REQ_PRSN
                                  , A.ERP_APPLY_FG
                                  , A.ERP_APPLY_ERROR
                                  FROM T_IF_SND_PUR_REQ_KO441 A (NOLOCK)
                                  WHERE 1=1
                                AND GW_APP_NO <> ''
                            ");

                if (DocumentNo.Length > 0)
                    sb.AppendLine(string.Format("AND A.GW_APP_NO LIKE '%{0}%'", DocumentNo));

                if (ReqPrsn.Length > 0)
                    sb.AppendLine(string.Format("AND A.REQ_PRSN  = (SELECT USR_ID FROM Z_USR_MAST_REC WHERE USR_NM LIKE '%{0}%')", ReqPrsn));

                if (PlantCd.Length > 0)
                    sb.AppendLine(string.Format("AND A.PLANT_CD = '{0}'", PlantCd));

                if (FromDay.Length > 0)
                    sb.AppendLine(string.Format("AND A.INSRT_DT >= '{0}'", FromDay));

                if (ToDay.Length > 0)
                    sb.AppendLine(string.Format("AND A.INSRT_DT <= '{0}'", ToDay));

                if (FlagType != "A")
                    sb.AppendLine(string.Format("AND A.ERP_APPLY_FG = '{0}'", FlagType));
                else
                    sb.AppendLine("AND A.ERP_APPLY_FG <> 'N'");
            }
            return sb.ToString();
        }
        #endregion

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            DataTable dtReturn = new DataTable();
            SqlDataAdapter Sqladapter;

            string sqlQuery = GetSelectQuery();

            try
            {
                SqlStateCheck();
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

        protected void dgList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataRowView rowView = (DataRowView)e.Row.DataItem;
            string[] tempArray;

            if (rowView == null)
                cnt = 0;
            else
            {
                tempArray = new string[rowView.Row.ItemArray.Length];

                for (int i = 0; i < tempArray.Length + 1; i++)
                {

                    if (i == 0)
                    {
                        if (rowView["ERP_APPLY_FG"].ToString() != "Y")
                            e.Row.Cells[0].Text = string.Format("<input type=\"button\" value=\"수정\" onclick=\"PopOpenUpdate('{0}')\" />", rowView["PR_NO"].ToString().Trim());
                        else
                            e.Row.Cells[0].Text = "-";

                        if ((cnt % 2) == 0) //가독성을 위한 짝수 Row는 색상 변경
                            e.Row.BackColor = Color.FromName("#D5D5D5");
                    }
                    else
                        e.Row.Cells[i].Text = rowView[i - 1].ToString();
                }

                cnt += 1;
            }
        }

        protected void dgList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            dgList.PageIndex = e.NewPageIndex;
            btnSelect_Click(sender, e);
        }

        //하단부에 있는 엑셀 업로드관련 공부하고 추가하기
        //protected void btnExcel_Click(object sender, EventArgs e)
        //{
        //    Response.Clear();
        //    //파일이름 설정
        //    string fName = string.Format("{0}.xls", DateTime.Now.ToString("yyyyMMddHHmmss"));
        //    //헤더부분에 내용을 추가
        //    Response.AddHeader("Content-Disposition", "attachment;filename=" + fName);
        //    Response.Charset = "utf-8";
        //    //컨텐츠 타입 설정
        //    string encoding = Request.ContentEncoding.HeaderName;
        //    Response.ContentType = "application/ms-excel";
        //    Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=" + encoding + "'>");

        //    System.IO.StringWriter SW = new System.IO.StringWriter();
        //    HtmlTextWriter HW = new HtmlTextWriter(SW);
        //    SW.WriteLine(" "); //한글 깨짐 방지

        //    dgList.RenderControl(HW);
        //    Response.Write(SW.ToString());
        //    Response.End();
        //    HW.Close();
        //    SW.Close();
        //}

        //public override void VerifyRenderingInServerForm(System.Web.UI.Control control)
        //{
        //    // Confirms that an HtmlForm control is rendered for the specified ASP.NET server control at run time.
        //}


    }
}