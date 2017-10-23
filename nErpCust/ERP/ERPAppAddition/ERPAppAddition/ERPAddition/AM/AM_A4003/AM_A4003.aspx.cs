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
using System.Drawing;
using System.IO;

namespace ERPAppAddition.ERPAddition.AM.AM_A4003
{
    public partial class AM_A4003 : System.Web.UI.Page
    {
        //속도개선 단가 없는 화면이 필요하다고 해서 복사해서 만들어줬음
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        string date = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                /*사이트 오픈 이력*/
                WebSiteCount();
                /*파라미터셋팅*/
                if (Request.QueryString["DATE"] == null || Request.QueryString["DATE"] == "")
                    date = "2016-07-01";
                else
                    date = Request.QueryString["DATE"];

                Session["date"] = date;


                /*조회*/
                retrieve();
                /*엑셀떨구기*/
                Excel_Click();
                /*창닫기*/
                string close = @"<script type='text/javascript'>
                                            window.returnValue = true;
                                            window.open('','_self','');                                                            
                                            window.close();
                                            </script>";
                base.Response.Write(close);
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void retrieve()
        {   
            // 프로시져 실행: 기본데이타 생성
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = setSql();
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

        private string setSql()
        {
            StringBuilder sbSQL = new StringBuilder();            
            string pDate = Session["date"].ToString();
            sbSQL.Append("EXEC USP_BANKACCOUNTLIST_VIEW  '" + pDate + "' \n");
            return sbSQL.ToString();
        }


        protected void Excel_Click()
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