using Microsoft.Reporting.WinForms;
using OrderDraftMailSending;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PurOderMailing
{


    public partial class Form1 : Form
    {
        public enum LogStatus : int
        {
            Info = 1
            ,
            Success = 2
                , Error = 3

        }
        private string Db;
        public Form1(string[] args)
        {
            InitializeComponent();

            Db = "";
            if (args.Length > 0)
            {
                if (args[0].Equals("test"))
                {
                    Db = "Nepes_DB_Test";
                }
                else if (args[0].Equals("display"))
                {
                    Db = "Nepes_Display";
                }
                else if (args[0].Equals("led"))
                {
                    Db = "Nepes_LED";
                }
                else
                {
                    Db = "Nepes_DB";
                }
            }
          

            if (Db.Equals(""))
            {
                return;
            }
        }


        private static string _url = System.Configuration.ConfigurationSettings.AppSettings["WebURL"];
        private static string _mailing_Path = System.Configuration.ConfigurationSettings.AppSettings["MAIL_PATH"];

        private static string _dbInfo;
        private static string _mailPostid = System.Configuration.ConfigurationSettings.AppSettings["MAIL_ADMIN"];
        private static string _saveLogPath = System.Configuration.ConfigurationSettings.AppSettings["LOG_PATH"];

        private static string _sDate;
        private static string _sFileName;
        private static string _sPath = System.Configuration.ConfigurationSettings.AppSettings["PDF_PATH"];

        static DBConn sqlConn = null;
        static LogControl logCtrl = null;

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime dt = new DateTime();

            logCtrl = new LogControl(_saveLogPath, Db);

            _dbInfo = System.Configuration.ConfigurationSettings.AppSettings[Db];

            dt = DateTime.Now;

            _sDate = dt.Year.ToString() + dt.Month.ToString("00") + dt.Day.ToString("00");
            //_sPath = "..\\PDF";
            _sPath = System.IO.Path.Combine(_sPath, _sDate);


            DataTable dtPoList = GetPurOderList();

            if (dtPoList.Rows.Count < 1)
            {
                logCtrl.IOFileWrite("발주 대상 없음", (int)LogStatus.Info);
                this.Close();
            }
            else
            {
                for (int i = 0; i < dtPoList.Rows.Count; i++)
                {

                    try
                    {
                        string sTitle = "";

                        if(Db == "Nepes_Display")
                        {
                            sTitle = "[nepes Display]발주서 송부  : [";
                        }
                        else if(Db == "Nepes_Display")
                        {
                            sTitle = "[nepes LED]발주서 송부  : [";
                        }
                        else
                        {
                            sTitle = "[nepes]발주서 송부  : [";
                        }

                        logCtrl.IOFileWrite("[" + dtPoList.Rows[i]["PO_NO"].ToString() + "] 작업 시작", (int)LogStatus.Info);
                        DataSet ds = SetPurOder(dtPoList.Rows[i]["PO_NO"].ToString());

                        string snd_yn = ds.Tables["PARAM:1"].Rows[0][0].ToString();
                        string sMailFrom = ds.Tables["Result"].Rows[0]["PO_EMAIL"].ToString();
                        string sMailTo = dtPoList.Rows[i]["TO_MAIL_ADDR"].ToString();
                        string sMailToGrp = dtPoList.Rows[i]["TO_MAIL_ADDR_GROUP"].ToString();
                        string sMailCC = dtPoList.Rows[i]["CC_MAIL_ADDR"].ToString();
                        string sSubject = sTitle + dtPoList.Rows[i]["BP_NM"].ToString() + "]" + dtPoList.Rows[i]["PO_NO"].ToString();
                        string sMsg_Body = dtPoList.Rows[i]["MAIL_BODY_CD"].ToString();


                        if (snd_yn == "Y")
                        {

                            GetPDFFile(ds.Tables["Result"]);

                            if(sMailFrom.Equals(""))
                            {
                                sMailFrom = "50439@nepes.co.kr";
                            }

                            if (SendingMail(sMailFrom, sMailTo, sMailToGrp, sMailCC, sSubject, sMsg_Body, false))
                            {
                                SetSendYN(dtPoList.Rows[i]["PO_NO"].ToString());
                                logCtrl.IOFileWrite("[" + dtPoList.Rows[i]["PO_NO"].ToString() + "] 작업 종료", (int)LogStatus.Success);
                            }
                            else
                            {
                                logCtrl.IOFileWrite("발주 메일 발송 실패[" + dtPoList.Rows[i]["PO_NO"].ToString() + "]", (int)LogStatus.Error);
                                this.Close();
                            }

                        }
                        else
                        {
                            string sErrMsg = ds.Tables["Result"].Rows[0]["ERROR_MSG"].ToString();
                            SetError(dtPoList.Rows[i]["PO_NO"].ToString(), "발주서 생성오류");
                            GetPDFFile(ds.Tables["Result"]);
                            if (SendingMail(sMailFrom, sMailTo, sMailToGrp, sMailCC, "[발주 오류]" + sSubject, sErrMsg, true))
                            {
                                
                            }
                            else
                            {
                                logCtrl.IOFileWrite("발주 오류 메일 발송 실패[" + dtPoList.Rows[i]["PO_NO"].ToString() + "]", (int)LogStatus.Error);
                                //this.Close();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SendingMail("50458@nepes.co.kr", "jsh0811@nepes.co.kr", "", "", "[발주 오류]" + dtPoList.Rows[i]["PO_NO"].ToString(), ex.Message, true);
                        logCtrl.IOFileWrite(" [" + _sFileName + "] 오류 발생 [" + ex.ToString() + "]", (int)LogStatus.Error);
                        SetError(dtPoList.Rows[i]["PO_NO"].ToString(), ex.Message.Replace("'","''"));
                        //this.Close();
                    }
                }
                this.Close();
            }

        }

        private DataSet SetPurOder(string Po)
        {
            DataSet ds = new DataSet();
            string sMailAddr = "";
            this.reportViewer1.RefreshReport();
            reportViewer1.Reset();

            try
            {
                if (sqlConn == null)
                    sqlConn = new DBConn(_dbInfo);

                List<SqlParameter> param = new List<SqlParameter>();

                SqlParameter param0 = new SqlParameter("@po_no", SqlDbType.VarChar, 30);
                SqlParameter param1 = new SqlParameter("@SND_YN", SqlDbType.VarChar, 1);

                param0.Value = Po;
                param1.Direction = ParameterDirection.Output;

                param.Add(param0);
                param.Add(param1);

                ds = sqlConn.ExecutePROCEDURE("dbo.USP_PO_VIEW_INFO", param);

                string snd_yn = ds.Tables["PARAM:1"].Rows[0][0].ToString();
                //if (snd_yn == "Y")
                {

                }

                return ds;
            }
            catch(Exception ex)
            {
                logCtrl.IOFileWrite(" [" + Po + "] PO 정보 조회 ERROR" + ex.Message, (int)LogStatus.Error);
                throw ex;
                //return ds;
            }

            
        }

        // sub 레포트 데이타 연결
        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {

            // conn_erp.Open();
            DataSet ds = new DataSet();
            List<SqlParameter> param = new List<SqlParameter>();
            SqlParameter param0 = new SqlParameter("@po_no", SqlDbType.VarChar, 30);

            param0.Value = _sFileName;

            param.Add(param0);

            ds = sqlConn.ExecutePROCEDURE("dbo.USP_PO_VIEW_DTL", param);

            try
            {

                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = ds.Tables[0];

                e.DataSources.Add(new ReportDataSource("DataSet1", ds.Tables[0]));
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(" [" + _sFileName + "] PDF 파일 DTL 생성 ERROR" + ex.Message, (int)LogStatus.Error);
                throw ex;

            }
        }


        private DataTable GetPurOderList()
        {
            DataTable dt = new DataTable();
            StringBuilder sSQL = new StringBuilder();

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("SELECT A.PO_NO");
            sSQL.AppendLine("     , B.TO_MAIL_ADDR");
            sSQL.AppendLine("     , (SELECT UD_MINOR_MSG_VALUE FROM B_USER_DEFINED_MINOR");
            sSQL.AppendLine("         WHERE UD_MAJOR_CD = 'M0012' ");
            sSQL.AppendLine("         AND UD_MINOR_CD = B.TO_MAIL_ADDR_GROUP) AS TO_MAIL_ADDR_GROUP");
            sSQL.AppendLine("     , B.CC_MAIL_ADDR");
            sSQL.AppendLine("     , B.TERMS_CD");
            sSQL.AppendLine("     , (SELECT BP_FULL_NM FROM B_BIZ_PARTNER WHERE BP_CD = A.BP_CD) AS BP_NM");
            sSQL.AppendLine("     , (SELECT UD_MINOR_MSG_VALUE FROM B_USER_DEFINED_MINOR");
            sSQL.AppendLine("         WHERE UD_MAJOR_CD = 'M0013' ");
            sSQL.AppendLine("         AND UD_MINOR_CD = B.MAIL_BODY_CD) AS MAIL_BODY_CD");
            sSQL.AppendLine("	FROM M_PUR_ORD_HDR A WITH (NOLOCK)");
            sSQL.AppendLine("INNER JOIN M_PUR_ORD_MAIL B WITH (NOLOCK)");
            sSQL.AppendLine(" ON A.PO_NO = B.PO_NO");
            sSQL.AppendLine(" WHERE 1=1");
            sSQL.AppendLine("  AND B.SEND_YN = 'N'");
            sSQL.AppendLine("  AND A.EXT1_CD = 'Y'");

            try
            {
                dt = sqlConn.ExecuteQuery(sSQL.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message, (int)LogStatus.Error);
            }

            return dt;

        }


        private void SetSendYN(string sPO)
        {
            StringBuilder sSQL = new StringBuilder();

            logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 완료 처리 START", (int)LogStatus.Info);

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("UPDATE M_PUR_ORD_MAIL");
            sSQL.AppendLine("   SET SEND_YN = 'Y'");
            sSQL.AppendLine("   , UPDT_DT = GETDATE()");
            sSQL.AppendLine(" WHERE 1=1");
            sSQL.AppendLine("  AND PO_NO = '" + sPO + "'");

            try
            {
                sqlConn.ExecuteQuery(sSQL.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 완료 처리 ERROR" + ex.Message, (int)LogStatus.Error);
            }
            logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 완료 처리 END", (int)LogStatus.Info);

        }

        private void SetError(string sPO, string sErrMSG)
        {
            StringBuilder sSQL = new StringBuilder();

            logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 오류 처리 START", (int)LogStatus.Info);

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("UPDATE M_PUR_ORD_MAIL");
            sSQL.AppendLine("   SET SEND_YN = 'E'");
            sSQL.AppendLine("   , ERR_MSG = '" + sErrMSG + "'");
            sSQL.AppendLine("   , UPDT_DT = GETDATE()");
            sSQL.AppendLine(" WHERE 1=1");
            sSQL.AppendLine("  AND PO_NO = '" + sPO + "'");

            try
            {
                sqlConn.ExecuteQuery(sSQL.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 오류 처리 ERROR" + ex.Message, (int)LogStatus.Error);
            }
            logCtrl.IOFileWrite(" [" + sPO + "] PO 메일발송 오류 처리 END", (int)LogStatus.Info);

        }

        private void GetPDFFile(DataTable dt)
        {

            try
            {
                if (dt.Rows.Count > 0)
                {


                    reportViewer1.Reset();

                    _sFileName = dt.Rows[0]["PO_NO"].ToString();

                    logCtrl.IOFileWrite(" [" + _sFileName + "] PDF 파일 생성 START", (int)LogStatus.Info);
                    


                    dt.Columns.Add("Company_NM");

                    if (Db == "Nepes_Display")
                    {
                        dt.Rows[0]["Company_NM"] = "Nepes Display";

                        _sPath = System.IO.Path.Combine(_sPath, "Display");
                    }
                    else if (Db == "Nepes_LED")
                    {
                        dt.Rows[0]["Company_NM"] = "Nepes LED";
                        _sPath = System.IO.Path.Combine(_sPath, "LED");
                    }
                    else
                    {
                        dt.Rows[0]["Company_NM"] = "Nepes Corporation";
                        _sPath = System.IO.Path.Combine(_sPath, "Semi");
                    }

                    logCtrl.IOFileWrite(" [" + _sPath + "] PDF 파일 위치", (int)LogStatus.Info);

                    reportViewer1.LocalReport.ReportPath = "rv_mm_m5001.rdlc";

                    reportViewer1.LocalReport.DisplayName = "발주서" + _sDate;

                    ReportDataSource rds = new ReportDataSource();
                    reportViewer1.ProcessingMode = ProcessingMode.Local;

                    rds.Name = "DataSet1";
                    rds.Value = dt;

                    reportViewer1.LocalReport.DataSources.Add(rds);

                    reportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);



                    this.reportViewer1.RefreshReport();

                }
                Warning[] warnings;
                string[] streamids;
                string mimeType;
                string encoding;
                string filenameExtension;

                reportViewer1.ProcessingMode = ProcessingMode.Local;
                byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);

                // using (FileStream fs = new FileStream("..\\PDF\\"+_sDate +"\\"+_sFileName+".PDF", FileMode.Create))
                System.IO.Directory.CreateDirectory(_sPath);
                using (FileStream fs = new FileStream(_sPath + "\\" + _sFileName + ".PDF", FileMode.Create))
                {
                    fs.Write(bytes, 0, bytes.Length);
                }
                logCtrl.IOFileWrite(" [" + _sFileName + "] PDF 파일 생성 END", (int)LogStatus.Info);
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(" [" + _sFileName + "] PDF 파일 생성 실패" + ex.Message, (int)LogStatus.Error);
                throw ex;
            }
        }

        private bool SendingMail(string msg_frm, string msg_to, string msg_to_grp, string msg_cc, string msg_subject, string msg_body, bool ErrYn)
        {
            System.Net.Mail.SmtpClient smtp = null;
            System.Net.Mail.MailMessage mailing = null;
            System.Net.Mail.Attachment attachment;

            bool sendMail = true;
            logCtrl.IOFileWrite(" [" + _sFileName + "] 메일 발송 시작", (int)LogStatus.Info);
            if (msg_to.Trim().Length == 0 && msg_to_grp.Trim().Length == 0)
            {
                sendMail = false;
                return sendMail;
            }

            try
            {
                mailing = new System.Net.Mail.MailMessage();
                mailing.From = new System.Net.Mail.MailAddress(msg_frm);

                if (Db.Equals("Nepes_DB_Test"))
                {
                    mailing.To.Add("jsh0811@nepes.co.kr");  // Test 모드일때 구매그룹 담당자로 메일 전송하지 않음
                }
                else
                {
                    mailing.To.Add(msg_frm);
                }
                if (ErrYn)
                {
                    //발주서 정보 오류시 발송 처리
                    //mailing.To.Add(msg_frm);
                    mailing.To.Add("jsh0811@nepes.co.kr");
                }
                else
                {
                    msg_to = msg_to.Replace(';', ',');
                    msg_to_grp = msg_to_grp.Replace(';', ',');

                    string[] srTO = msg_to.Trim().Split(',');
                    string[] srTOGrp = msg_to_grp.Trim().Split(',');

                    if (srTO.Length > 0)
                    {
                        //받는 사람
                        for (int i = 0; i < srTO.Length; i++)
                        {
                            if (!srTO[i].Trim().Equals(""))
                            {
                                mailing.To.Add(srTO[i].Trim());
                            }
                        }
                    }
                    if (srTOGrp.Length > 0)
                    {
                        //받는 사람 그룹 => 그룹 받는 사람에서 참조 그룹으로 변경 
                        for (int i = 0; i < srTOGrp.Length; i++)
                        {
                            if (!srTOGrp[i].Trim().Equals(""))
                            {
                                mailing.CC.Add(srTOGrp[i].Trim());
                            }
                        }
                    }

                    //참조
                    if (msg_cc.Length > 0)
                    {
                        msg_cc = msg_cc.Replace(';', ',');
                        string[] srCC = msg_cc.Trim().Split(',');

                        if (srCC.Length > 0)
                        {
                            for (int c = 0; c < srCC.Length; c++)
                            {
                                mailing.CC.Add(srCC[c]);
                            }
                        }
                    }

                }

                mailing.Bcc.Add("jsh0811@nepes.co.kr"); //개발자 메일모니터링을 위해 숨은참조를 추가, 안정화 이후 주석처리 예정 
                mailing.Priority = System.Net.Mail.MailPriority.High;
                mailing.Subject = msg_subject;
                mailing.Body = msg_body;
                mailing.IsBodyHtml = false;
                mailing.DeliveryNotificationOptions = System.Net.Mail.DeliveryNotificationOptions.OnSuccess;

                attachment = new System.Net.Mail.Attachment(_sPath + "\\" + _sFileName + ".PDF");

                mailing.Attachments.Add(attachment);
                smtp = new System.Net.Mail.SmtpClient(System.Configuration.ConfigurationSettings.AppSettings["MAIL_HOST"]);

                smtp.Send(mailing);

                logCtrl.IOFileWrite(" [" + _sFileName + "] 메일 발송 완료", (int)LogStatus.Info);
                return sendMail;
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(" [" + _sFileName + "] 메일 발송 실패 : "+ ex.Message, (int)LogStatus.Error);
                throw ex;

            }
            finally
            {
                if (mailing != null) mailing = null;
                if (smtp != null) smtp = null;
            }
        }

    }
}
