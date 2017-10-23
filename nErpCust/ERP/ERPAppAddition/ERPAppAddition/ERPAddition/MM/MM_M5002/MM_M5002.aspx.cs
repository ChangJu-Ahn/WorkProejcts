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
using System.Drawing;
using System.Drawing.Imaging;
using ERPAppAddition.QueryExe;
using Microsoft.Reporting.WebForms;
using FarPoint.Web.Spread.Data;
using System.Net;
using System.Web.Mail;

namespace ERPAppAddition.ERPAddition.MM.MM_M5002
{
    public partial class MM_M5002 : System.Web.UI.Page
    {
        bool sendOk = true;
        bool fileOk = true;
        StringBuilder filePath;
        bool fileExist = false;
        StringBuilder errorMessage = new StringBuilder("");

        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_display"].ConnectionString);
        SqlCommand cmd_erp = new SqlCommand();
        SqlDataReader dr_erp;
        SqlDataAdapter erp_sqlAdapter;
        DataSet ds = new DataSet();

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;
        string sql_spread;
        int value, chk_save_yn = 0;
        string userid, db_name;
        cls_dbexe_erp dbexe = new cls_dbexe_erp();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {

                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "") //사용자 ID값이 없다면 개발자 ID로할지 판단하기
                {
                    if (Request.QueryString["db"] == null || Request.QueryString["db"] == "") //DB없이 바로 실행할때 개발자용으로 적용
                        userid = "dev"; //erp에서 실행하지 않았을시 대비용
                    else // DB명이 있는데 사용자 ID가 없다면 이상하니 다시 접속하라는 메세지 보여줌
                    {
                        MessageBox.ShowMessage("잘못된 접근입니다. ERP접속 후 실행해주세요", this.Page);
                        this.Response.Redirect("../../Fail_Page.aspx");
                    }
                }

                else
                    userid = Request.QueryString["userid"];

                Session["User"] = userid;
                WebSiteCount();
            }

        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_erp.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void GetImage(string idx) //이미지 가져오기
        {
            if (Request.QueryString["idx"] == null)
            {
                Response.End();
            }
            else
            {
                GetImage(Request.QueryString["idx"]);
            }

            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.usp_po_getimage";
            cmd_erp.CommandTimeout = 3000;
            SqlParameter spIdx = new SqlParameter("@idx", SqlDbType.Int);
            spIdx.Value = idx;

            dr_erp = cmd_erp.ExecuteReader();

            while (dr_erp.Read())
            {
                byte[] image = (byte[])dr_erp["Image"];
                MemoryStream ms = new MemoryStream(image, 0, image.Length);
                Bitmap bitmap = new Bitmap(ms);
                System.Drawing.Image im = System.Drawing.Image.FromStream(ms);
                Response.ContentEncoding = System.Text.Encoding.UTF8;


                Response.ContentType = "image/jpeg/png";
                Response.AddHeader("Content-Disposition", "attachment; filename="
                   + Server.UrlEncode(dr_erp["FileName"].ToString()));

                bitmap.Save(Response.OutputStream, ImageFormat.Jpeg);
            }

            dr_erp.Close();
            dr_erp.Close();
            dr_erp.Dispose();
            dr_erp.Dispose();
            dr_erp.Dispose();
        }





        // sub 레포트 데이타 연결
        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {

            // conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.StoredProcedure;
            cmd_erp.CommandText = "dbo.USP_PO_VIEW_DTL";
            cmd_erp.CommandTimeout = 3000;
            SqlParameter param1 = new SqlParameter("@PO_NO", SqlDbType.VarChar, 30);

            param1.Value = tb_po_no.Text;
            cmd_erp.Parameters.Add(param1);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                DataTable dt = new DataTable();
                da.Fill(dt);

                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = dt;

                e.DataSources.Add(new ReportDataSource("DataSet1", dt));
            }
            catch { }
            finally
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
        }



        protected void btn_pop_po_Click1(object sender, EventArgs e) //발주번호 찾기 팝업창
        {
            Response.Write("<script>window.open('pop_mm_m5002_po.aspx?pgid=mm_m5002&popupid=2','','top=100,left=100,width=1000,height=600')</script>");

        }

        protected void rbt_select_SelectedIndexChanged(object sender, EventArgs e) //추가입력,메일 선택 기능
        {
            if (rbt_select.SelectedValue == "keyin") //추가입력 선택
            {
                Panel_keyin.Visible = true;
              

            }

              

            if (rbt_select.SelectedValue == "update_cp") //Contact Person 선택
            {

                Panel_keyin.Visible = false;
              

            }
        }
     



        private void Delete_File()
        {
            try
            {
                FileInfo myFile = new FileInfo(filePath.ToString());
                myFile.Delete();
            }
            catch (Exception e)
            {
            }
        }

       

        protected void btn_keyin_save_Click(object sender, EventArgs e) //추가입력 key-in 저장버튼
        {
            conn_erp.Open();

            string sql = "insert into m_po_view_keyin(po_no, Warranty, remark) " +
                        "values('" + tb_po_no.Text + "','" + tb_keyin_warranty.Text + "','" + tb_keyin_remark.Text + "')";


            SqlCommand sComm = new SqlCommand(sql, conn_erp);
            MessageBox.ShowMessage("저장되었습니다.", this.Page);

            sComm.ExecuteNonQuery();
            conn_erp.Close();
        }


        protected void btn_update_Click(object sender, EventArgs e) //추가입력 key-in 수정버튼
        {
            conn_erp.Open();

            string sql = "update m_po_view_keyin set Warranty = '" + tb_keyin_warranty.Text + "' , remark ='" + tb_keyin_remark.Text + "' " +
                         " where po_no='" + tb_po_no.Text + "'";


            SqlCommand sComm = new SqlCommand(sql, conn_erp);
            MessageBox.ShowMessage("수정되었습니다.", this.Page);

            sComm.ExecuteNonQuery();
            conn_erp.Close();
        }




        protected void btn_retrieve_Click1(object sender, EventArgs e)
        {

            string po_no = tb_po_no.Text;
            if (po_no == null || po_no == "" || tb_po_no.Text.Equals(""))
            {
                MessageBox.ShowMessage(".", this.Page);
                return;
            }


            else
            {
                ReportViewer1.Reset();
                conn_erp.Open();
                cmd_erp = conn_erp.CreateCommand();
                cmd_erp.CommandType = CommandType.StoredProcedure;
                cmd_erp.CommandText = "dbo.USP_PO_VIEW";
                cmd_erp.CommandTimeout = 3000;
                SqlParameter param1 = new SqlParameter("@po_no", SqlDbType.VarChar, 30);

                param1.Value = tb_po_no.Text;
                cmd_erp.Parameters.Add(param1);


                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    ReportViewer1.LocalReport.ReportPath = Server.MapPath("rv_mm_m5002.rdlc");
                    ReportViewer1.LocalReport.DisplayName = tb_po_no.Text + "발주서" + DateTime.Now.ToShortDateString();
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "DataSet1";
                    rds.Value = dt;
                    ReportViewer1.LocalReport.DataSources.Add(rds);
                    ReportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
                    ReportViewer1.LocalReport.Refresh();

                    UpdatePanel1.Update();


                }
                catch (Exception ex)
                {
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                }
            }

        }

        public string attach1 { get; set; }

        public object attach { get; set; }

       


    }
}
