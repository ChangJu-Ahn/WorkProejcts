using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ERPAppAddition.ERPAddition.SM.sm_sa001;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;

namespace ERPAppAddition.ERPAddition.SM.sm_sb001
{
    public partial class sm_sb001_A01_TEST : System.Web.UI.Page
    {
        sa_fun fun = new sa_fun();        

        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["NEPES_MAIL_DEV"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();


        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;

                //MessageBox.ShowMessage(userid, this);

                /*공장*/
                DataTable plant = fun.getData("SELECT PLANT_DESC, GROUP1_CODE FROM DBO.SA_SYS_CODE_EKP order by 2");
                if (plant.Rows.Count > 0)
                {
                    DDL_PLANT.DataTextField = "PLANT_DESC";
                    DDL_PLANT.DataValueField = "GROUP1_CODE";
                    DDL_PLANT.DataSource = plant;
                    DDL_PLANT.DataBind();
                }
                DateTime setDate = DateTime.Today.AddDays(0);
                tb_fr_yyyymmdd.Text = setDate.Year.ToString("0000") + setDate.Month.ToString("00") + setDate.Day.ToString("00");
                tb_to_yyyymmdd.Text = DateTime.Today.Year.ToString("0000") + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");

                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Button button1 = (Button)sender;
            GridViewRow grdrow = (GridViewRow)button1.Parent.Parent;
            string PROCESS_INSTANCE_OID = grdrow.Cells[7].Text;
            string uid = Session["User"].ToString();

            Response.Write("<script>window.name='main'</script>");
            /*운영*/
            //Response.Write("<script>window.open('sm_sb001_A02.aspx?process_instance_oid=" + PROCESS_INSTANCE_OID + "&userid=" + uid + "','','resizable=no, top=100,left=100,width=1000,height=650')</script>");
            /*개발*/
            Response.Write("<script>window.open('sm_sb001_A02_TEST.aspx?process_instance_oid=" + PROCESS_INSTANCE_OID + "&userid=" + uid + "','','resizable=no, top=100,left=100,width=1000,height=650')</script>");            
        }

        protected void btn_select(object sender, EventArgs e)
        {
            search();
        }

        private void search()
        {
            DataSet ds = new DataSet();
            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    DataRow dr = ds.Tables["DataSet1"].NewRow();
                    ds.Tables["DataSet1"].Rows.InsertAt(dr, 0);

                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);

                    //dt.Rows.Add(new object[] { "" });
                    GridView1.Columns[6].Visible = false;
                    //return;
                }
                else
                {
                    GridView1.Columns[6].Visible = true;
                }

                GridView1.DataSource = ds.Tables["DataSet1"];
                GridView1.DataBind();
                
                
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
        }
        private string getSQL()
        {
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            string strFac = DDL_PLANT.Text;
            string strDOC_NO = DOC_NO.Text;
            string strCREATOR = CREATOR.Text;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("USP_DIST_11  '" + strFrom + "', '" + strTo + "', '" + strFac + "', '" + strDOC_NO + "', '" + strCREATOR + "' \n");
            return sbSQL.ToString();

        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //StringBuilder builder = new StringBuilder();
            //string strFileName = "반출정보조회" + DateTime.Now.ToShortDateString() + ".xls";
            //builder.Append("Name ,Education,Location" + Environment.NewLine);
            //foreach (GridViewRow row in GridView1.Rows)
            //{
            //    string name = row.Cells[0].Text;
            //    string education = row.Cells[1].Text;
            //    string location = row.Cells[2].Text;
            //    builder.Append(name + "," + education + "," + location + Environment.NewLine);
            //}
            //Response.Clear();
            //Response.ContentType = "application/vnd.ms-excel"; 
            //Response.AddHeader("Content-Disposition", "attachment;filename=" + strFileName);
            //Response.Write(builder.ToString());
            //Response.Charset = "euc-kr";
            //Response.ContentEncoding = System.Text.Encoding.GetEncoding("euc-kr"); 
            //Response.End();

            //Response.Clear();
            //Response.AddHeader("Content-Disposition", "attachment;filename=data.xls");
            //Response.Charset = "";

            //Response.ContentType = "application/vnd.xml";
            //Response.CacheControl = "public";

            //System.IO.StringWriter strinWrite = new System.IO.StringWriter();
            //System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(strinWrite);
            //GridView1.RenderControl(htmlWrite);
            //Response.Write(strinWrite.ToString());
            //Response.End();


            Response.Clear();
            //파일이름 설정
            string fName = string.Format("WorkCode{0}.xls", DateTime.Now.ToString("yyyyMMddHHmmss"));
            //헤더부분에 내용을 추가
            Response.AddHeader("Content-Disposition", "attachment;filename=" + fName);
            Response.Charset = "utf-8";
            //컨텐츠 타입 설정
            string encoding = Request.ContentEncoding.HeaderName;
            Response.ContentType = "application/unknown";
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