using System.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Web;
using System.Resources;
using System.IO;

using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace OrderDraftMailSending
{
    public partial class Program
    {

        #region 멤버변수 선언(mailsend, url, mailing_Part, dbInfo, mailPostid)
        private static string _url = System.Configuration.ConfigurationSettings.AppSettings["WebURL"];
        private static string _mailing_Path = System.Configuration.ConfigurationSettings.AppSettings["MAIL_PATH"];
        private static string _dbInfo = System.Configuration.ConfigurationSettings.AppSettings["Nepes_DB"];
        private static string _mailPostid = System.Configuration.ConfigurationSettings.AppSettings["MAIL_ADMIN"];
        private static string _saveLogPath= System.Configuration.ConfigurationSettings.AppSettings["LOG_PATH"];
        #endregion

        #region db연결을 위한 클래스객체 생성(DBConn) / 로그저장을 위한 클래스 객체 생성(LogControl)
        static OrderDraftMailSending.DBConn sqlConn = null;
        static OrderDraftMailSending.LogControl logCtrl = new LogControl(_saveLogPath);
        #endregion

        static void Main()
        {
            SendMail();
        }

        private static void SendMail()
        {
            DataTable dtList = null; //보내야할 메일정보를 담는 데이터테이블

            dtList = GetSendMailList("PO");

            if (dtList.Rows.Count > 0 && !dtList.Equals(null))
            {
               int cnt = dtList.Rows.Count;            //개별로 메일을 보내야하기 때문에 로우 카운트

                for (int i = 0; i < cnt; i++)
                {
                    DataRow drHeader;
                    DataRow drDetail;
                    string strERP_No = string.Empty;
                    string strSend_Mail = string.Empty;

                    strERP_No = dtList.Rows[i]["PO_NO"].ToString().Trim();         //헤더, 디테일 select에 사용 될 PO번호 대입
                    strSend_Mail = dtList.Rows[i]["SEND_EMAIL"].ToString().Trim(); //이메일 보낸사람 (구매그룹에 설정되어 있는 이메일주소)

                    drHeader = GetPoHeaderDate(strERP_No);
                    drDetail = GetPoDetailDate(strERP_No);

                    if (!drHeader.Equals(null) && !drDetail.Equals(null))
                    {
                        HtmlParsing(drHeader, drDetail, strSend_Mail); //데이터테이블의 정보를 가지고 html 파싱
                        //ErpFlagUpdate(strERP_No);        //메일을 보냈기 때문에 ERP FLAG 업데이트
                    }
               }
            }
        }

        private static DataRow GetPoHeaderDate(string po_no)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                DataRow drResult;

                if (sqlConn == null)
                    sqlConn = new DBConn(_dbInfo);

                sb.AppendLine(" SELECT * FROM ufn_po_document_serach('" + po_no + "')");
                drResult = sqlConn.ExecuteQuery(sb.ToString()).Rows[0];

                return drResult;
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message);
                return null;
            }
        }

        private static DataRow GetPoDetailDate(string po_no)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                DataRow drResult;

                if (sqlConn == null)
                    sqlConn = new DBConn(_dbInfo);

                sb.AppendLine(" USP_PO_SENDMAIL_DETAIL '" + po_no + "'");
                drResult = sqlConn.ExecuteQuery(sb.ToString()).Rows[0];

                return drResult;
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message);
                return null;
            }
        }

        private static DataTable GetSendMailList(string type)
        {   
            try
            {
                StringBuilder sb = new StringBuilder();
                DataTable dtResult = new DataTable();

                if (sqlConn == null)
                    sqlConn = new DBConn(_dbInfo);

                switch (type.ToUpper()) 
                {  
                    case "PO" :
                        sb.AppendLine(" SELECT * FROM ");
                        sb.AppendLine(" T_IF_SND_PUR_POMAIL_KO441(nolock) WHERE SEND_FLAG = 'N' ");
                        dtResult = sqlConn.ExecuteQuery(sb.ToString());
                        break;

                    default :
                        dtResult = null;
                        break;
                }

                return dtResult;
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message);
                return null;
            }
        }


        private static bool SendSMTP(string msg_frm, string msg_to, string msg_cc, string msg_subject, string msg_body)
        {
            System.Net.Mail.SmtpClient smtp = null;
            System.Net.Mail.MailMessage mailing = null;

            try
            {
                mailing = new System.Net.Mail.MailMessage();

                mailing.From = new System.Net.Mail.MailAddress(msg_frm);
                mailing.To.Add(msg_to);
                mailing.CC.Add(msg_cc);
                mailing.Bcc.Add("ahncj@nepes.co.kr,kimjr@nepes.co.kr"); //2016.06.08, ahncj : 개발자 메일모니터링을 위해 숨은참조를 추가, 안정화 이후 주석처리 예정 
                mailing.Priority = System.Net.Mail.MailPriority.High;
                mailing.Subject = msg_subject;
                mailing.Body = msg_body;
                mailing.IsBodyHtml = true;

                smtp = new System.Net.Mail.SmtpClient(System.Configuration.ConfigurationSettings.AppSettings["MAIL_HOST"]);

                smtp.Send(mailing);

                return true;
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (mailing != null) mailing = null;
                if (smtp != null) smtp = null;
            }
        }

        private static void HtmlParsing(DataRow drHeader, DataRow drDetail, string Send_Mail)
        {
            Hashtable hashdt = new Hashtable(); //매개변수로 넘겨받는 데이터로우를 해시테이블에 값, 키로 구분하여 저장
            
            string htmlDetailBody = string.Empty;
            string strFilePath = string.Empty;
            string strHTML = string.Empty;
            string htmlTemp = string.Empty;
            string logPath = string.Empty;
            string logContent = string.Empty;
            string MailSubject = string.Empty;
            bool mailState = false;

            try 
            { 
                hashdt.Add("DAY", Convert.ToDateTime(drHeader["PO_DAY"]).ToString("yyyy-MM-dd"));
                hashdt.Add("GW_NO", drHeader["GW_PO_NO"].ToString());
                hashdt.Add("PLANT", drHeader["PLANT_ADDRESS"].ToString());
                hashdt.Add("PR_PERSON", drHeader["PR_PERSON_NM"].ToString());
                hashdt.Add("PR_TELNO", drHeader["PR_PERSON_TEL"].ToString());
                hashdt.Add("PO_DEPT", drHeader["PO_DEPT"].ToString());
                hashdt.Add("PARTNER_NM", drHeader["PARTNER_NM"].ToString());
                hashdt.Add("PO_PERSON", drHeader["PO_PERSON"].ToString());
                hashdt.Add("PO_PERSON_TITLE", drHeader["PO_PERSON_TITLE"].ToString());
                hashdt.Add("PARTNER_PERSON_NM", drHeader["PARTNER_PERSON_NM"].ToString());
                hashdt.Add("PO_TEL", drHeader["PO_TEL"].ToString());
                hashdt.Add("PO_FAX", drHeader["PO_FAX"].ToString());
                hashdt.Add("PO_HP_TEL", drHeader["PO_HP_TEL"].ToString());
                hashdt.Add("PO_EMAIL", drHeader["PO_EMAIL"].ToString());
                hashdt.Add("PARTNER_TEL", drHeader["PARTNER_TEL"].ToString());
                hashdt.Add("PARTNER_FAX", drHeader["PARTNER_FAX"].ToString());
                hashdt.Add("PARTNER_HP_TEL", "-");
                hashdt.Add("PARTNER_EMAIL", drHeader["PARTNER_EMAIL"].ToString());

                htmlDetailBody = drDetail[0].ToString();

                strFilePath = _mailing_Path + "mail.htm";                           //메일 html을 불러오기 위한 구문
                StreamReader sr = new StreamReader(strFilePath, Encoding.Default);  //html 인코딩 시 한글깨짐 관련 수정(한글깨짐 관련 코딩추가)        
                strHTML = sr.ReadToEnd();
                sr.Close();

                MailSubject = string.Format("[nepes]발주서 송부의 건({0})", System.DateTime.Now.ToString("yyyy-MM-dd"));

                //헤더정보 html 파싱
                htmlTemp = strHTML.Replace("#orderDay", hashdt["DAY"].ToString());
                htmlTemp = htmlTemp.Replace("#order_Gw_No", hashdt["GW_NO"].ToString());
                htmlTemp = htmlTemp.Replace("#plant", hashdt["PLANT"].ToString());
                htmlTemp = htmlTemp.Replace("#pr_Person", hashdt["PR_PERSON"].ToString());
                htmlTemp = htmlTemp.Replace("#pr_Telnumber", hashdt["PR_TELNO"].ToString());
                htmlTemp = htmlTemp.Replace("#po_Dept", hashdt["PO_DEPT"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_Name", hashdt["PARTNER_NM"].ToString());
                htmlTemp = htmlTemp.Replace("#po_Person", hashdt["PO_PERSON"].ToString());
                htmlTemp = htmlTemp.Replace("#Person_Title", hashdt["PO_PERSON_TITLE"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_Person", hashdt["PARTNER_PERSON_NM"].ToString());
                htmlTemp = htmlTemp.Replace("#po_Tel", hashdt["PO_TEL"].ToString());
                htmlTemp = htmlTemp.Replace("#po_Fax", hashdt["PO_FAX"].ToString());
                htmlTemp = htmlTemp.Replace("#po_HP_Tel", hashdt["PO_HP_TEL"].ToString());
                htmlTemp = htmlTemp.Replace("#po_Email", hashdt["PO_EMAIL"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_Tel", hashdt["PARTNER_TEL"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_Fax", hashdt["PARTNER_FAX"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_HP_Tel", hashdt["PARTNER_HP_TEL"].ToString());
                htmlTemp = htmlTemp.Replace("#partner_Email", hashdt["PARTNER_EMAIL"].ToString());

                //디테일정보 html 파싱
                htmlTemp = htmlTemp.Replace("#body", htmlDetailBody);
            

                //메일 보내기 전 로그 
                if (!logCtrl.LogPathString.Equals("") && !logCtrl.LogPathString.Equals(null))
                    LogWrite(drHeader["PO_NO"].ToString(), drHeader["GW_PO_NO"].ToString(), "전송 전");      //로그메시지 생성  

                //파싱된 html을 메일로 전송
                //SendSMTP(_mailPostid, "ahncj@nepes.co.kr,kimjr@nepes.co.kr,abc@abc.com", "ahncj@nepes.co.kr", "발주서테스트입니다", htmlTemp);
                mailState = SendSMTP(Send_Mail, hashdt["PARTNER_EMAIL"].ToString(), Send_Mail, MailSubject, htmlTemp);

                if (mailState == true)
                    ErpFlagUpdate(drHeader["PO_NO"].ToString());        //메일을 보냈기 때문에 ERP FLAG 업데이트

                if (!logCtrl.LogPathString.Equals("") && !logCtrl.LogPathString.Equals(null))
                    LogWrite(drHeader["PO_NO"].ToString(), drHeader["GW_PO_NO"].ToString(), "전송 완료");      //로그메시지 생성  
                
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message);
            }
        }

        private static void ErpFlagUpdate(string Po_no)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                if (sqlConn == null)
                    sqlConn = new DBConn(_dbInfo);

                sb.AppendLine(" UPDATE T_IF_SND_PUR_POMAIL_KO441 ");
                sb.AppendLine("   SET SEND_DT = GETDATE(), SEND_FLAG = 'Y' ");
                sb.AppendLine(" WHERE PO_NO = '" + Po_no + "' ");

                sqlConn.ExecuteNonQuery(sb.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message);
            }
        }

        private static void LogWrite(string po_no, string gw_po_no, string msg)
        {
            string logMsg = string.Empty;

            logMsg = string.Format("PO_NO : {0}, GW_NO : {1}, ({2})", po_no, gw_po_no, msg);      //로그메시지 생성  
            logCtrl.IOFileWrite(logMsg);    //로그작성
        }

    }
}
