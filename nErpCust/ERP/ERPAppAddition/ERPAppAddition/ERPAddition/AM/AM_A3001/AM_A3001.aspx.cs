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
using ERPAppAddition.ERPAddition.AM.AM_A3001;


namespace ERPAppAddition.ERPAddition.AM.AM_A3001
{
    public partial class AM_A30011 : System.Web.UI.Page
    {

        SqlConnection conn ;//= new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_display"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;

        string ls_gl_no;
        string userid, db_name;
        protected void Page_Load(object sender, EventArgs e)
        {
            
            //userid = Request.QueryString["userid"];
            //string db_name = String.Empty;
            if (!Page.IsPostBack)
            {
                if (Request.QueryString["db"] != null && Request.QueryString["db"].ToString() != "")
                {
                    db_name = Request.QueryString["db"].ToString();
                    if (db_name.Length > 0)
                    {
                        conn = new SqlConnection(ConfigurationManager.ConnectionStrings[db_name].ConnectionString);
                    }
                    userid = Request.QueryString["userid"];

                    Session["DBNM"] = Request.QueryString["db"].ToString();
                    Session["User"] = Request.QueryString["userid"];
                }
                else
                {
                    string script = "alert(\"프로그램 호출이 잘못되었습니다. 관리자에게 연락해주세요.\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                }
                
                FillGrid();
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

        private void FillGrid()
        {
           
            DataTable dt = fetch();
            GridView1.DataSource = dt;
            GridView1.DataBind();
        }

        public DataTable fetch()
        {
           // userid = "songth";
           // db_name = "nepes_test1";
          //  string script = "alert(\"userid=" + userid + "db_name=" + db_name + "\");";
          //  ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[Session["DBNM"].ToString()].ConnectionString);
            string sql = "  select temp_gl_no /*결의전표번호*/   " +
                         "       , convert(varchar(10),TEMP_GL_DT,102)  TEMP_GL_DT /*결의일자*/ " +
	                     "       , (select sdeptnm " +
	                     "            from HORG_MAS x inner join HORG_ABS y on x.orgid = y.orgid " +
		                 "           where x.DEPT = a.dept_cd " +
		                 "             and y.ORGDT = (select max(ORGDT) " +
		                 "                              from HORG_ABS  " +
						 "                             where a.TEMP_GL_DT >= convert(datetime,orgdt) ) " +
                         " 	       ) dept_nm /*부서*/ " +
                         "       , DR_AMT /*금액)*/ " +
	                     "       , DR_LOC_AMT /*금액(자국)*/ " +
                         "       , INSRT_USER_ID    " + 
                         "   from a_temp_gl a " +
                         "  where  a.TEMP_GL_DT between '" + tb_fr_dt.Text + "' and '" + tb_to_dt.Text + "' and INSRT_USER_ID = '" + Session["User"].ToString() + "'  " +
                         "    and  ( DIST_TYPE NOT IN ('Y','E') OR DIST_TYPE IS NULL) " +
                         "    and  CONF_FG <> 'C' " +
                         "  order by 1 " ;
            AM_A3001 ds = new AM_A3001();
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);

            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            conn.Close();
            return ds.Tables[0];

        }

        protected void btn_exe_Click(object sender, EventArgs e)
        {
            
            DataTable dt = fetch();
            GridView1.DataSource = dt;
            GridView1.DataBind();
        }

        protected void btn_gw_Click(object sender, EventArgs e)
        {
             // //선택된 드랍다운리스트가 있는 row찾기
            GridView grid = sender as GridView;
            
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                RadioButton rdol_value = (RadioButton)GridView1.Rows[i].FindControl("rbtn_check");

                if (rdol_value.Checked == true)
                {
                    ls_gl_no = GridView1.DataKeys[i].Values[0].ToString(); //기준년월일

                    string url = "http://mail.nepes.co.kr/Logon/Legacy/WFLinkage.aspx?LEGACY_CODE=L_StatementRecv&ERP_KEY=" + ls_gl_no + "&GL_NO=" + ls_gl_no + "&USER_ID=" + Session["User"].ToString() + "&ERP_DB=" + Session["DBNM"].ToString();
                    //string script = "alert(\"" + url +  "\");";
                    //ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    //string url = "http://gw.devnepes.co.kr/Logon/Legacy/WFLinkage.aspx?LEGACY_CODE=L_Statement&ERP_KEY=" + ls_gl_no + "&GL_NO=" + ls_gl_no + "&USER_ID=" + Session["User"].ToString() + "&ERP_DB=" + Session["DBNM"].ToString();
                    //Response.Redirect(url, true);


                    //string Fullurl = "newpage.aspx/";
                    OpenNewBrowserWindow(url, this);
                    btn_exe_Click(null, null);
                }
            }
            
        }

        public static void OpenNewBrowserWindow(string Url, Control control)
        {
            ScriptManager.RegisterStartupScript(control, control.GetType(), "Open", "window.open('" + Url + "');", true);
        }
       
    }
}