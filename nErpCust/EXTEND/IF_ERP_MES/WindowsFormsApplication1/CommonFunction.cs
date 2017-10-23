using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace T_IF_RCV_PROD_ORD_KO441
{
    public static class CommonFunction
    {
        #region SendMail
        public static bool SendMail(string sendAddress, string sendUser, string toUser, string ccUser, string subject, string bodyConent, string sendFileListFullPath, string sendFileName, bool bBodyHtml, bool bFileExist)
        {
            try
            {
                string strMailData = string.Empty;

                // send
                MailMessage mailMsg = new MailMessage();
                mailMsg.Headers.Clear();
                mailMsg.From = new MailAddress(sendAddress, sendUser);
                mailMsg.To.Add(toUser);
                //mailMsg.CC.Add(ccUser); 참조는 필요 없으니 주석처리 (실제 "" 으로 매개변수를 전달 받았으나 공란이라 애러가 발생되어 주석처리 함)

                mailMsg.Subject = subject;
                mailMsg.Body = bodyConent;
                mailMsg.BodyEncoding = System.Text.Encoding.UTF8;
                mailMsg.IsBodyHtml = bBodyHtml;

                if (bFileExist == true)
                {
                    Attachment attFile = new Attachment(sendFileListFullPath);
                    attFile.Name = sendFileName;
                    mailMsg.Attachments.Add(attFile);
                }

                System.Net.NetworkCredential crMail = new System.Net.NetworkCredential();
                //crMail.UserName = "ahncj@nepes.co.kr"; // 개인메일을 사용하였으나 공통계정으로 변경 함
                //crMail.Password = "1234"; // 개인메일을 사용하였으나 공통계정으로 변경 함
                crMail.UserName = "nepes_sys@nepes.co.kr";
                crMail.Password = "nepes123";

                SmtpClient sClient = new SmtpClient();

                sClient.Host = "mail.nepes.co.kr";
                sClient.Port = 25;
                sClient.UseDefaultCredentials = false;
                sClient.Credentials = crMail;

                sClient.Send(mailMsg);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        #endregion
    }
}
