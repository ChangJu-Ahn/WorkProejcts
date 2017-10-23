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
using ERPAppAddition.ERPAddition.IM.IM_I1001;
using System.Drawing;
using System.IO;

namespace ERPAppAddition.ERPAddition.IM.IM_I1001
{
    public partial class IM_I1002 : System.Web.UI.Page
    {
        //속도개선
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        cls_prod_qty_month cls_dbexe = new cls_prod_qty_month();

        SqlDataReader dr = null;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
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

        protected void bt_retrieve_Click(object sender, EventArgs e)
        {
            string ITEM_CD = "";
            if (tb_item_cd.Text == "" || tb_item_cd.Text == null)
                ITEM_CD = "%";
            else
                ITEM_CD = tb_item_cd.Text;

            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "EXEC USP_STOCK_AGE_EXT_1 '" + ddl_plant_cd.Text + "', '" + ITEM_CD + "'";
            cmd.CommandTimeout = 3000;

            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                Console.WriteLine("{0} Second exception caught.", ex);
            }
            GridView1.DataSource = dt;
            GridView1.DataBind();
         
        }

        protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowIndex >= 0)
            {
                for (int i = 6; i < e.Row.Cells.Count; i++)
                {
                    if (e.Row.Cells[i].Text != "&nbsp;")
                    {
                        if (Convert.ToDouble(e.Row.Cells[i].Text) > 0)
                        {
                            if (i <= 11)
                            {
                                e.Row.Cells[i].CssClass = "grpCell1";

                            }
                            else if (i <= 13)
                            {
                                e.Row.Cells[i].CssClass = "grpCell2";
                            }
                            else if (i <= 15)
                            {
                                e.Row.Cells[i].CssClass = "grpCell3";
                            }
                            else
                            {
                                e.Row.Cells[i].CssClass = "grpCell4";
                            }
                        }
                        else
                        {
                            e.Row.Cells[i].CssClass = "grpCellW";
                        }
                    }
                }
            }
        }

        protected void GridView1_RowCreated(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.Header)
            {
                GridView oGridView = (GridView)sender;
                GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);

                TableCell oTableCell = new TableCell();
                oTableCell.Text = "";
                oTableCell.ColumnSpan = 6;
                oGridViewRow.Cells.Add(oTableCell);

                string[] hdrAdd = { "≤15", "16 - 30", "31 - 60", "61 - 90", "91 - 120", "121 - 150", "151 - 180", "181 - 210", "  211 - 240", "241 - 270", "≥ 271" };

                foreach (string str in hdrAdd)
                {
                    TableCell TableCell = new TableCell();
                    TableCell.Text = str;
                    TableCell.ColumnSpan = 2;
                    TableCell.HorizontalAlign = HorizontalAlign.Center;
                    TableCell.BackColor = Color.CornflowerBlue;
                    oGridViewRow.Cells.Add(TableCell);
                }
                oGridView.Controls[0].Controls.AddAt(0, oGridViewRow);
            }
        }

        protected void Excel_Click(object sender, EventArgs e)
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

            GridView1.RenderControl(HW);
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