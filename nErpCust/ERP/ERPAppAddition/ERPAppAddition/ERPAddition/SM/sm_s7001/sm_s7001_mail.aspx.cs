using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Net.Mail;


namespace ERPAppAddition.ERPAddition.SM.sm_s7001
{
    public partial class sm_s7001_mail : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void Capture(object sender, EventArgs e)
        {
            startCaputre();
        }

        private void startCaputre()
        {
            string url = "http://192.168.10.98:369/ERPAddition/SM/sm_s7001/sm_s7001_sub.aspx";
            Thread thread = new Thread(delegate()
            {
                using (WebBrowser browser = new WebBrowser())
                {
                    browser.ScrollBarsEnabled = false;
                    browser.AllowNavigation = true;
                    browser.Navigate(url);
                    browser.Width = 1024;
                    browser.Height = 768;
                    //browser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(DocumentCompleted);//웹페이지 로딩이 끝나고 실행될 이벤트

                    
                    while (browser.ReadyState != WebBrowserReadyState.Complete || browser.Document.Body.ScrollRectangle.Height < 500 )
                    {
                        System.Windows.Forms.Application.DoEvents();//이벤트 실행                                                
                    }
                    
                    if(browser.ReadyState == WebBrowserReadyState.Complete )
                    {                        
                        DocumentCompleted(browser);
                    }
                    
                }
            });
            thread.SetApartmentState(ApartmentState.STA);//STA : 싱글 쓰레드, MTA : 멀티 쓰레드
            thread.Start();
            thread.Join();
        }
        

        private void DocumentCompleted(object sender)
        {
            WebBrowser browser = sender as WebBrowser;

            browser.Height = browser.Document.Body.ScrollRectangle.Height + 30;
            browser.Width = browser.Document.Body.ScrollRectangle.Width;


            using (Bitmap bitmap = new Bitmap(browser.Width, browser.Height))
            {
                browser.DrawToBitmap(bitmap, new Rectangle(0, 0, browser.Width, browser.Height));

                Bitmap SaveImage = new Bitmap(browser.Width, browser.Height);                

                int ReSizeX = 0;
                int ReSizeY = 0;

                string Path = Server.MapPath(".") + @"\a.jpg";
                string Path2 = Server.MapPath(".") + @"\a2.jpg";


                Graphics gr = Graphics.FromImage(SaveImage);
                gr.DrawImage(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height), new Rectangle(ReSizeX, ReSizeY, bitmap.Width - ReSizeX, bitmap.Height - ReSizeY), GraphicsUnit.Pixel);// 이 부분에서 위치 조정 가능
                gr.Dispose();

                SaveImage.Save(Path2);
                //SaveImage.Save(Path, System.Drawing.Imaging.ImageFormat.Jpeg);
                
                browser.Dispose();

                string html = "<html><body>";
                html += "<div>";
                html += "<img src=\"cid:img\">";
                html += "</div>";
                html += "</body></html>";

                AlternateView altView = AlternateView.CreateAlternateViewFromString(html, null, System.Net.Mime.MediaTypeNames.Text.Html);

                LinkedResource theEmailImage = new LinkedResource(Path2);
                theEmailImage.ContentId = "img";

                altView.LinkedResources.Add(theEmailImage);

                //MailMessage msg = new MailMessage("yoosr@nepes.co.kr", "yoosr@nepes.co.kr");

                MailMessage msg = new MailMessage();

                msg.To.Add("janghs0501@nepes.co.kr");
                //msg.CC.Add("ahncj@nepes.co.kr");
                msg.Subject = "음성 매출현황";
                msg.IsBodyHtml = true;
                msg.AlternateViews.Add(altView);



                //var sClient = new SmtpClient("mail.nepes.co.kr", 25);
                var sClient = new SmtpClient();
                sClient.UseDefaultCredentials = false;
                using (sClient as IDisposable)
                {
                    sClient.Send(msg);
                }
                msg.Dispose();
                txtUrl.Text = Path;

//                string close = @"<script type='text/javascript'>
//                                window.returnValue = true;
//                                window.open('','_self','');                                                            
//                                window.close();
//                                </script>";
//                base.Response.Write(close);

            }

        }
    }

}