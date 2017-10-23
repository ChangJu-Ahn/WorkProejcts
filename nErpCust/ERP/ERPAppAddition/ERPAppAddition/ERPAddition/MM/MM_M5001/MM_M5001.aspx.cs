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




namespace ERPAppAddition.ERPAddition.MM.MM_M5001
{
    public partial class MM_M5001 : System.Web.UI.Page
    {


            bool sendOk=true; 
            bool fileOk=true; 
            StringBuilder filePath; 
            bool fileExist=false; 
            StringBuilder errorMessage = new StringBuilder(""); 

        SqlConnection conn_erp = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
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
            Response.Write("<script>window.open('pop_mm_m5001_po.aspx?pgid=mm_m5001&popupid=1','','top=100,left=100,width=1000,height=600')</script>");

        }

        protected void rbt_select_SelectedIndexChanged(object sender, EventArgs e) //추가입력,메일 선택 기능
        {
            if (rbt_select.SelectedValue == "keyin") //추가입력 선택
            {
                Panel_keyin.Visible = true;
                Panel_mail.Visible = false;
                
            }
          
            if (rbt_select.SelectedValue == "mail") //메일 선택
            {
                Panel_keyin.Visible = false;
                Panel_mail.Visible = true;
                
            }

            if (rbt_select.SelectedValue == "update_cp") //Contact Person 선택
            {
           
                Panel_keyin.Visible = false;
                Panel_mail.Visible = false;
                
            }
        }
        


        protected void btn_send_Click1(object sender, EventArgs e) //메일전송
        {
            Check_Message();
            if (!sendOk)
            {
                Response.Write("<script language='javascript'>"
                    + "alert(‍'" + errorMessage.ToString() + "')"
                    + "<" + "/script>");
            }
            else
            {
                Upload_File();
                if (fileOk)
                {
                    MailMessage Message1 = new MailMessage();
                    Message1.To = txt_mail_to.Text;
                    Message1.Cc = txt_mail_cc.Text;
                    Message1.From = txt_mail_fr.Text;
                    Message1.Subject = txt_mail_subject.Text;
                    Message1.Body = txt_mail_message.Text;
                   // Message1.BodyFormat = MailFormat.Html;
                    Message1.BodyEncoding = Encoding.Default;
                    if (fileExist)
                    {
                        MailAttachment Attach1 = new MailAttachment(filePath.ToString(), MailEncoding.Base64);
                        Message1.Attachments.Add(Attach1);
                    }


                    Response.Write("<script language='javascript'>"
                    + "window.status='메일을 보내는 중...'"
                    + "<" + "/script>");

                    SmtpMail.SmtpServer = "mail.nepes.co.kr" ; 
                    SmtpMail.Send(Message1);


                    Response.Write("<!--\n\n\n-->");
                    Response.Write("<script>");
                    Response.Write("alert(\"메일을 발송하였습니다.\\n" +
                                    "From : " + this.txt_mail_fr.Text.ToString() +
                                    "\\nTo : " + this.txt_mail_to.Text.ToString() +
                                    "\\nCC : " + this.txt_mail_cc.Text.ToString() + "\");");

                    Response.Write("</script>");

                    Delete_File();
                }
                else
                {
                    Response.Write("<script language='javascript'>"
                        + "alert(‍(\"파일 송신 중에 에러가 생겼습니다..\\n");
                    Response.Write("</script>");
                }
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

        private void Upload_File()
        {
            if (FileUpload1.PostedFile.FileName != "")
            {
                try
                {
                    FileInfo upFile = new FileInfo(FileUpload1.PostedFile.FileName);
                    filePath = new StringBuilder(Server.MapPath(upFile.Name));
                    FileInfo saveFile;
                    do
                    {
                        saveFile = new FileInfo(filePath.ToString());
                        if (saveFile.Exists)
                        {
                            filePath = filePath.Replace(".", "1.");
                        }
                    }
                    while (saveFile.Exists);
                    FileUpload1.PostedFile.SaveAs(filePath.ToString());
                }
                catch (Exception e)
                {
                    fileOk = false;
                }
                fileExist = true;
            }
        }

        private void Check_Message()
        {
            if (txt_mail_to.Text == "")
            {
                sendOk = false;
                errorMessage.Append("- 받는 사람을 적어 주세요.\\n\\n");
            }
            if (txt_mail_fr.Text == "")
            {
                sendOk = false;
                errorMessage.Append("- 보내는 사람을 적어 주세요.\\n\\n");
            }
            if (txt_mail_subject.Text == "")
            {
                sendOk = false;
                errorMessage.Append("- 제목을 적어 주세요.\\n\\n");
            }
            if (txt_mail_message.Text  == "")
            {
                sendOk = false;
                errorMessage.Append("- 내용을 적어 주세요.\\n\\n");
            } 
        }
 
            //MailMessage msg = new MailMessage();

            //msg.From = new MailAddress(txt_mail_fr.Text); //보내는 사람 메일

            //if (!"".Equals(this.txt_mail_to.Text.ToString())) //받는 사람 메일
            //{
            //    MailAddressCollection mc_to = getMailAddrCollection(this.txt_mail_to.Text.ToString());
            //    for (int i = 0; i < mc_to.Count; i++)
            //        msg.To.Add(mc_to[i].Address);
            //}

            //if (!"".Equals(this.txt_mail_cc.Text.ToString())) //참조
            //{
            //    MailAddressCollection mc_cc = getMailAddrCollection(this.txt_mail_cc.Text.ToString());
            //    for (int i = 0; i < mc_cc.Count; i++)
            //        msg.CC.Add(mc_cc[i].Address);
            //}


            //msg.Subject = this.txt_mail_subject.Text.ToString(); //제목
            //msg.Body = this.txt_mail_message.Text.ToString();    //내용

            ////----------------------------------------첨부파일 넣기--------------------------------------------------------

            //            string attachFileName1 = ""; //첨부파일추가1
            //            string attachFileName2 = ""; //첨부파일추가2


            //            if (this.FileUpload1.HasFile)
            //            {
            //                attachFileName1 = saveAttachFile(FileUpload1, attachFileName1);
            //                if (!"".Equals(attachFileName1))
            //                {
            //                    FileUpload1.PostedFile.SaveAs(attachFileName1);
            //                    msg.Attachments.Add(new Attachment(attachFileName1));
            //                }
            //            }
            //            if (this.FileUpload2.HasFile)
            //            {
            //                attachFileName2 = saveAttachFile(FileUpload2, attachFileName2);
            //                if (!"".Equals(attachFileName2))
            //                {
            //                    FileUpload2.PostedFile.SaveAs(attachFileName2);
            //                    msg.Attachments.Add(new Attachment(attachFileName2));
            //                }
            //            }

            //            client.UseDefaultCredentials = true;
            //            client.Send(msg);


            //            Response.Write("<!--\n\n\n-->");
            //            Response.Write("<script>");
            //            Response.Write("alert(\"메일을 발송하였습니다.\\n" +
            //                            "From : " + this.txt_mail_fr.Text.ToString() +
            //                            "\\nTo : " + this.txt_mail_to.Text.ToString() +
            //                            "\\nCC : " + this.txt_mail_cc.Text.ToString() + "\");");

            //            Response.Write("</script>");

            //        }


            //        protected string saveAttachFile(FileUpload fu, string uploadDir)//파일 첨부시 이름 같을 경우
            //        {
            //            string attachFileName = "";
            //            if (fu.PostedFile.InputStream.Length > 0)
            //            {
            //                string path = fu.PostedFile.FileName;
            //                string fileName = path.Substring(path.LastIndexOf("\\") + 1);
            //                fileName = getValidFileName(uploadDir, fileName);
            //                attachFileName = uploadDir + "\\" + fileName;
            //            }

            //            return attachFileName;
            //        }


            //        protected string getValidFileName(string uploadDir, string fileName) //유효한 파일명으로 변경
            //        {
            //            if (!File.Exists(uploadDir + "\\" + fileName))
            //                return fileName;

            //            string ext = fileName.Substring(fileName.LastIndexOf('.'));
            //            fileName = fileName.Substring(0, fileName.LastIndexOf('.'));

            //            int i = 0;
            //            while (File.Exists(uploadDir + "\\" + fileName + "[" + i + "]" + ext))
            //            {
            //                i++;
            //            }

            //            return fileName + "[" + i + "]" + ext;



            //        }


            //        protected MailAddressCollection getMailAddrCollection(string str) //여러메일 동시 전송용
            //        {
            //            MailAddressCollection mac = new MailAddressCollection();

            //            if (!"".Equals(str))
            //            {
            //                string[] arr_str = str.Split(';');
            //                for (int i = 0; i < arr_str.Length; i++)
            //                {
            //                    if (!"".Equals(arr_str[i].Trim()))
            //                    {
            //                        MailAddress ma = new MailAddress(arr_str[i].Trim());
            //                        mac.Add(ma);
            //                    }
            //                }
            //            }

            //            return mac;
            //        }


        protected void btn_pop_mail_Click(object sender, EventArgs e) // 메일 보내는 사람 주소록 팝업창
        {
            Response.Write("<script>window.open('pop_mm_m5001_mail.aspx?pgid=mm_m5001&popupid=1','','top=100,left=100,width=500,height=500')</script>");

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
            string msg = "";
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
                //cmd_erp.CommandText = "dbo.USP_PO_VIEW"; 기존
                cmd_erp.CommandText = "dbo.USP_PO_VIEW_NEW"; // 151023, ahncj (신규 발주서 프로시저 적용)
                cmd_erp.CommandTimeout = 3000;
                SqlParameter param1 = new SqlParameter("@po_no", SqlDbType.VarChar, 30);
                SqlParameter param2 = new SqlParameter("@ERROR", SqlDbType.VarChar, 100);

                param1.Value = tb_po_no.Text;
                param2.Direction = ParameterDirection.Output; // 151023, ahncj (output 변수를 받아오기 위한 선언)
                cmd_erp.Parameters.Add(param1);
                cmd_erp.Parameters.Add(param2);

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd_erp);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    msg = (string)param2.Value; // 151023, ahncj (실제 output 변수로 넘어온 값을 msg 라는 변수에 대입)

                    if(dt.Rows.Count > 0 ) // 151022, ahncj (실제 발주번호가 없을 경우 경고문을 출력하기 위함)
                    {
                        if (msg.ToString() != "정상")  // 151023, ahncj (정상으로 출력될 경우 output msg는 '정상'이므로, '정상'인 메시지를 제외하고는 경고창을 출력)
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert(' " + msg + "')", true);

                        ReportViewer1.LocalReport.ReportPath = Server.MapPath("rv_mm_m5001.rdlc");
                        ReportViewer1.LocalReport.DisplayName = tb_po_no.Text + "발주서" + DateTime.Now.ToShortDateString();
                        ReportDataSource rds = new ReportDataSource();
                        rds.Name = "DataSet1";
                        rds.Value = dt;
                        ReportViewer1.LocalReport.DataSources.Add(rds);
                        ReportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
                        ReportViewer1.LocalReport.Refresh();

                        UpdatePanel1.Update();
                    }
                    else // 151022, ahncj (실제 발주번호가 없을 경우 경고문을 출력하기 위함)
                    { 
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 발주번호에 대한 정보가 없습니다.')", true);
                    }


                }
                catch (Exception ex) // 151022, ahncj (오류가 발생될 경우 경고문 출력)
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 발주번호의 정보가 올바르지 않습니다. \\n관리자에게 문의하세요.')", true);
                    if (conn_erp.State == ConnectionState.Open)
                        conn_erp.Close();
                }
            }
           
        }

        public string attach1 { get; set; }

        public object attach { get; set; }

      
    }

   }






        