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

namespace ERPAppAddition.ERPAddition.TEMP
{
    public partial class TempStockMovementReport : System.Web.UI.Page
    {
        #region Global Variable Declaration (Sql Connection, Command, Reader)
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sqlCmd;
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    txtdate_From.Text = DateTime.Now.ToString("yyyyMMdd");
                    txtdate_To.Text = DateTime.Now.ToString("yyyyMMdd");
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("fn_OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                }
                finally
                {
                    if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                        sqlConn.Close();
                }
            }
            
        }

        protected void SqlState_Check()
        {
            if (sqlConn == null || sqlConn.State == ConnectionState.Closed) sqlConn.Open();
        }

        protected void btnSelect_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            GridView grid = new GridView();
            string sqlQuery = GetSelectQuery();

            try
            {
                SqlState_Check();

                sqlCmd = new SqlCommand(sqlQuery, sqlConn);
                sqlCmd.CommandTimeout = 4000;
                da.SelectCommand = sqlCmd;
                da.SelectCommand.CommandTimeout = 4000;
                da.Fill(dt);

                dgList.BorderStyle = BorderStyle.Inset;
                dgList.BorderColor = Color.Black;
                dgList.HeaderStyle.BorderStyle = BorderStyle.Inset;
                dgList.HeaderStyle.BorderColor = Color.Black;
                dgList.HeaderStyle.BackColor = Color.LightSteelBlue;
                dgList.HeaderStyle.Font.Bold = true;
                dgList.RowStyle.HorizontalAlign = HorizontalAlign.Center;
                dgList.RowStyle.VerticalAlign = VerticalAlign.Middle;
                dgList.DataSource = dt;

                if (dt.Rows.Count > 0) dgList.DataBind();
                else
                {
                    dgList.Controls.Clear();
                    ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "fn_alert('검색된 정보가 없습니다.');", true);
                }


                //ReportViewerSetting(dt);
            }
            catch (Exception ex)
            {
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", string.Format("fn_OutputAlert('{0}');", ex.Message.Replace("'", "").ToString()), true);
                //MessageBox.ShowMessage("아래 내용을 관리자에게 문의하세요.\\n * 내용 : [" + ex.Message.Replace("'", " ") + "]", Page); //문자에 작은따옴표가 들어가 있을 경우 스크립트 애러가 발생되므로 작은따옴표를 공백으로 처리
            }
            finally
            {
                if (sqlConn != null || sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
            }
        }

        protected string GetSelectQuery()
         {
            StringBuilder sb = new StringBuilder();
            string from_Day = txtdate_From.Text.ToString();
            string to_Day = txtdate_To.Text.ToString();
            //string from_Day = txtdate_From1.Value.ToString();
            //string to_Day = txtdate_To1.Value.ToString();
            string query_Item = (txtItem.Text.Length > 0) ? "AND MES_ITEM_CD LIKE '%" + txtItem.Text.Trim() + "%'" : "";

            sb.AppendLine(string.Format(@"
                                SELECT CONVERT(CHAR(10), ACTUAL_GI_DT, 120) AS '재고이동일'
	                                ,PLANT_CD AS '공장'
	                                ,dbo.ufn_GetERP_ITEM_CD(MES_ITEM_CD, PLANT_CD) AS 'ERP품목코드'
	                                ,MES_ITEM_CD AS 'MES품목코드'
	                                ,CASE 
			                                WHEN (
					                                SELECT COUNT(*)
					                                FROM B_ITEM_BY_PLANT
					                                WHERE PLANT_CD = 'P04'
						                                AND ITEM_CD = dbo.ufn_GetERP_ITEM_CD(MES_ITEM_CD, PLANT_CD)
					                                ) > 0
				                                THEN 'Y'
			                                ELSE 'N'
			                                END
		                             AS '공장별정보유무(ERP코드기준)'
	                                ,dbo.ufn_GetSlCd(PLANT_CD, [DBO].[ufn_GetERP_ITEM_CD](MES_ITEM_CD, PLANT_CD)) AS '창고'
	                                ,'*' AS '로트번호'
	                                ,GI_QTY AS '수량'
	                                ,CASE WHEN dbo.ufn_GetISSUED_UNIT(PLANT_CD, [DBO].[ufn_GetERP_ITEM_CD](MES_ITEM_CD, PLANT_CD)) <> 'EE' THEN dbo.ufn_GetISSUED_UNIT(PLANT_CD, [DBO].[ufn_GetERP_ITEM_CD](MES_ITEM_CD, PLANT_CD))
                                          ELSE '-'
                                     END AS '단위'
	                                ,'*' AS 이동로트번호
	                                ,PLANT_CD 이동공장
	                                ,'043010' AS 이동창고
	                                ,([DBO].[ufn_GetERP_ITEM_CD](MES_ITEM_CD, PLANT_CD)) AS 이동품목
                                FROM T_IF_RCV_PART_OUT_KO441(NOLOCK)
                                WHERE CREATE_TYPE = 'A'
	                                AND PLANT_CD = 'P04'
	                                AND [DBO].[ufn_GetKIND_PARTOUT](PLANT_CD, OUT_TYPE, SHIP_TO_PARTY, SHIP_TO_PARTY_LINE) IN ('TX1')
	                                AND [DBO].[ufn_GetIssueReqNo](REPLACE(REMARK, ' ', '')) = 'N'
	                                --AND isnull([DBO].[ufn_GetERP_ITEM_CD](MES_ITEM_CD, PLANT_CD), '') <> ''
	                                AND ERP_APPLY_FLAG = 'Y'
                                    AND CONVERT(CHAR(10), ACTUAL_GI_DT, 112) BETWEEN '{0}' AND '{1}'
                                    {2}
                                ORDER BY ACTUAL_GI_DT
            ", from_Day, to_Day, query_Item));
            return sb.ToString();
        }

        protected void btnExcelDown_Click(object sender, EventArgs e)
        {
            Response.Clear();
            //파일이름 설정
            string fName = string.Format("{0}.xls", DateTime.Now.ToString("yyyyMMddHHmmss"));
            //헤더부분에 내용을 추가
            Response.AddHeader("Content-Disposition", "attachment;filename=" + fName);
            Response.Charset = "utf-8";
            //컨텐츠 타입 설정
            string encoding = Request.ContentEncoding.HeaderName;
            Response.ContentType = "application/ms-excel";
            Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=" + encoding + "'>");

            System.IO.StringWriter SW = new System.IO.StringWriter();
            HtmlTextWriter HW = new HtmlTextWriter(SW);
            SW.WriteLine(" "); //한글 깨짐 방지

            dgList.RenderControl(HW);
            Response.Write(SW.ToString());
            Response.End();
            HW.Close();
            SW.Close();
        }

        public override void VerifyRenderingInServerForm(System.Web.UI.Control control)
        {
            // Confirms that an HtmlForm control is rendered for the specified ASP.NET server control at run time.
        }

    }
}