using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;
using System.Collections;
using nInterface;

namespace nInterface
{
    class CommonFunction : Form1
    {

        //it is Constructor
        public CommonFunction()
        {

        }

        /// <summary>
        /// 메일을 전송합니다.
        /// </summary>
        /// <param name="lstContents">전송되어야 하는 메시지내용 입니다 (list 형식)</param>
        /// <param name="hs">ini을 파싱 한 정보입니다. (메일주소, 휴대폰그룹 등)</param>
        public bool setSendMail(List<string> lstContents, Hashtable hs)
        {
            System.Net.Mail.SmtpClient smtp = null;
            System.Net.Mail.MailMessage mailing = null;
            int rowcnt = 1;
            StringBuilder sb = new StringBuilder();

            //메일을 보낼 내용 조합
            sb.AppendLine("<B>아래와 같은 I/F오류가 발생되었습니다.</B>");
            sb.AppendLine("<B>자세한 부분은 시간과 오류내용을 참고하여 로그파일을 확인 바랍니다.</B>");
            sb.AppendLine("<BR><BR>");
            sb.AppendLine("<table border='1' style=\"font-size: 11px;\">");
            sb.AppendLine("<tr><th> No </th><th> Time </th><th> Contents </th>");

            foreach (string Contents in lstContents)
            {
                sb.AppendLine(string.Format("<TR><TD>{0}.</TD><TD>{1}</TD><TD>{2}</TD></TR>", rowcnt, System.DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss':'ffffff"), Contents.ToString()));
                sb.AppendLine(string.Format("{0}", Environment.NewLine));

                rowcnt++;
            }

            sb.AppendLine("</table>");

            try
            {
                mailing = new System.Net.Mail.MailMessage();

                mailing.From = new System.Net.Mail.MailAddress("50458@nepes.co.kr");
                mailing.To.Add(hs["TARGET"].ToString());
                mailing.Priority = System.Net.Mail.MailPriority.High;
                mailing.Subject = "[nepes]INTERFACE 애러발생 알람메일";
                mailing.Body = sb.ToString();
                mailing.IsBodyHtml = true;

                smtp = new System.Net.Mail.SmtpClient(System.Configuration.ConfigurationSettings.AppSettings["MAIL_HOST"]);

                smtp.Send(mailing);

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (mailing != null) mailing = null;
                if (smtp != null) smtp = null;
            }
        }


        /// <summary>
        /// SMS를 전송합니다 (확인은 오라클 MIGHTY DB에서 "SELECT * FROM SMSLIST WHERE PLANT = 'ERP'" 으로 확인)
        /// </summary>
        /// <param name="lstContents">전송되어야 하는 메시지내용 입니다 (list 형식)</param>
        /// <param name="hs">ini을 파싱 한 정보입니다. (메일주소, 휴대폰그룹 등)</param>
        public bool setSendSMS(List<string> lstContents, Hashtable hs)
        {
            string tempQuery = string.Empty;
            string smsType = hs["TARGET"].ToString(); //해당 타입은 mighty db에 저장되어있는 그룹을 의미한다. 그룹 조회는 다음의 쿼리를 참고하자 SELECT * FROM SMSLIST WHERE PLANT = 'ERP'
            string dbPart = System.Configuration.ConfigurationSettings.AppSettings["DB_PATH"];

            try
            {
                //프로시저 조합(따로 파라메터를 조합하여 만들 수 있지만 동일한 ExcuteQuery 함수를 사용하기 위해 string으로 조합 후 사용
                //[파라메터를 사용하려면 SQLCOMMEND 및 SQLPARAMETER를 선언해야 하기 때문에 불필요한 코드가 발생 됨]
                //프로시저는 MSSQL의 프로시저를 사용하며 프로시저의 내용은 MSSQL에서 확인하자.
                tempQuery = "EXEC SEND_SMS 'ERP', '" + smsType + "', 'IF오류 발생(" + lstContents.Count.ToString() + "건), 아래시간을 참고하여 로그확인 요망'";
                
                //객체 선언
                nInterface.DBConn Dbc = new DBConn(System.IO.File.ReadAllText(dbPart + "ERPInfo.ini").ToString(), "ERP");
                
                //DB Open
                Dbc.OpenERPDBConn();
                Dbc.ExecuteQuery(tempQuery, "ERP");

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


    }
}
