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
    public partial class sm_sb001_A03 : System.Web.UI.Page
    {
        sa_fun safun = new sa_fun();
        sb_fun sbfun = new sb_fun();

        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        string userid;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["NEPES_MAIL"].ConnectionString);
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
                DataTable plant = safun.getData("SELECT PLANT_DESC, GROUP1_CODE FROM DBO.SA_SYS_CODE_EKP order by 2");
                if (plant.Rows.Count > 0)
                {
                    DDL_PLANT.DataTextField = "PLANT_DESC";
                    DDL_PLANT.DataValueField = "GROUP1_CODE";
                    DDL_PLANT.DataSource = plant;
                    DDL_PLANT.DataBind();
                }

                /*반출상태*/
                DataTable dtSts = sbfun.getData("SELECT CODE_NM FROM DBO.TB_CERTLOG_CODE WHERE CODE_GROUP = 'A001' ORDER BY VALUE5");
                if (dtSts.Rows.Count > 0)
                {
                    DDL_STS.DataTextField = "CODE_NM";
                    DDL_STS.DataValueField = "CODE_NM";
                    DDL_STS.DataSource = dtSts;
                    DDL_STS.DataBind();
                }

                /*반입상태*/
                DataTable dtRests = sbfun.getData("SELECT CODE_NM FROM DBO.TB_CERTLOG_CODE WHERE CODE_GROUP = 'A002' ORDER BY VALUE5");
                if (dtRests.Rows.Count > 0)
                {
                    DDL_RESTS.DataTextField = "CODE_NM";
                    DDL_RESTS.DataValueField = "CODE_NM";
                    DDL_RESTS.DataSource = dtRests;
                    DDL_RESTS.DataBind();
                }

                /*미완료사유*/
                DataTable dtIncom = sbfun.getData("SELECT CODE_NM FROM DBO.TB_CERTLOG_CODE WHERE CODE_GROUP = 'A003' ORDER BY VALUE5");
                if (dtIncom.Rows.Count > 0)
                {
                    DDL_INCOMP.DataTextField = "CODE_NM";
                    DDL_INCOMP.DataValueField = "CODE_NM";
                    DDL_INCOMP.DataSource = dtIncom;
                    DDL_INCOMP.DataBind();
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

        protected void btn_select(object sender, EventArgs e)
        {
            search();
            GridView1_DataBound();
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
                }
                GridView1.DataSource = ds.Tables["DataSet1"];
                GridView1.DataBind();

                if (ds.Tables["DataSet1"].Rows.Count > 0)
                {
                 
                }

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }
        }
        

        public void GridView1_DataBound()
        {
            for (int j = 0; j < 11; j++)
            {
                int RowSpan = 2;
                for (int i = GridView1.Rows.Count - 2; i >= 0; i--)
                {
                    GridViewRow currRow = GridView1.Rows[i];
                    GridViewRow prevRow = GridView1.Rows[i + 1];
                                        
                        if (currRow.Cells[j].Text == prevRow.Cells[j].Text)
                        {
                            if (j == 0)
                            {
                                currRow.Cells[j].RowSpan = RowSpan;
                                prevRow.Cells[j].Visible = false;
                                RowSpan += 1;
                            }

                            else if(!prevRow.Cells[j - 1].Visible)
                            {
                                currRow.Cells[j].RowSpan = RowSpan;
                                prevRow.Cells[j].Visible = false;
                                RowSpan += 1;
                            } 
                            else
                                RowSpan = 2;
                        }
                        else
                            RowSpan = 2;                    
                }
            }          
            
        }

        private string getSQL()
        {
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            string strFac = DDL_PLANT.Text;
            string strDOC_NO = DOC_NO.Text;
            string strCREATOR = CREATOR.Text;

            string strSTS = DDL_STS.Text;      //반출상태
            string strRESTS = DDL_RESTS.Text;  //반입상태
            string strGOODS_NM = DDL_GOODS_NM.Text;     //품명
            string strINCOMPLETE_NM = DDL_INCOMP.Text;  //미완료사유

            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("USP_DIST_15_REP  '" + strFrom + "', '" + strTo + "', '" + strFac + "', '" + strDOC_NO + "', '" + strCREATOR + "', '" + strSTS + "', '" + strRESTS + "', '" + strGOODS_NM + "', '" + strINCOMPLETE_NM + "' \n");
            return sbSQL.ToString();

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