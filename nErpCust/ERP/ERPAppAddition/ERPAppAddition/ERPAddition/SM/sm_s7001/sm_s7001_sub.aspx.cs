using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
//using System.Data.OleDb;
//using System.Data.OracleClient;
using System.IO;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Data.OleDb;
using System.Net.Mail;


namespace ERPAppAddition.ERPAddition.SM.sm_s7001
{
    public partial class sm_s7001_sub : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
                
        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        SqlCommand sql_cmd2 = new SqlCommand();
        SqlCommand sql_cmd3 = new SqlCommand();
        SqlCommand sql_cmd4 = new SqlCommand();
        SqlCommand sql_cmd5 = new SqlCommand();
                
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();        
        DataTable dtYYYYMM= new DataTable();
        DataTable ddt = new DataTable();
        string tb_yyyymm = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

            }
            DateTime dt = DateTime.Now.AddDays(-1);
            tb_yyyymm = dt.Year.ToString("0000") + dt.Month.ToString("00") + dt.Day.ToString("00");

            getData();
            ReportViewer1.LocalReport.Refresh();

            /*report view 파일떨구기*/
            Warning[] warnings;
            string[] streamids;

            ddt = ds.Tables["DataSet1"].DefaultView.ToTable(true, "ITEM_NM", "BP_NM");
            /* ROW * 1   // 파이 // 차트 */
            //ddt.Rows.Count * 0.630 + 16 + 11;
            double pageHeight = 24;

            string deviceInfo = "<DeviceInfo><PageWidth>38cm</PageWidth>" +
                                    "<PageHeight>" + pageHeight + "cm</PageHeight>" +
                                    "<MarginTop>0in</MarginTop>" +
                                    "<MarginLeft>0in</MarginLeft>" +
                                    "<MarginBottom>0in</MarginBottom>" +
                                    "<MarginRight>0in</MarginRight>" +
                                    "<OutputFormat>JPEG</OutputFormat></DeviceInfo>";


            string mimeType = string.Empty;
            string encoding = string.Empty;
            string extension = string.Empty;
            byte[] bytes = ReportViewer1.LocalReport.Render("Image", deviceInfo, out mimeType, out encoding, out extension, out streamids, out warnings);

            string Path = Server.MapPath(".") + @"\a.jpg";
            FileStream fs = new FileStream(Path, FileMode.Create, FileAccess.Write);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            string html = "<html>";
            html += " <head>                                                                              \n ";
            html += "     <meta http-equiv='Content-Type' content='text/html; charset=utf-8' />           \n ";
            html += "     <style>                                                                         \n ";
            html += "         table                                                                       \n ";
            html += "         {                                                                           \n ";
            html += "             font-family: 'Lucida Grande' , Helvetica, Arial, Verdana, sans-serif;   \n ";
            html += "             border-collapse: collapse;                                              \n ";
            html += "             border-left: 1px solid #000;                                            \n ";
            html += "             border-top: 1px solid #000;                                             \n ";
            html += "         }                                                                           \n ";
            html += "         table thead tr th                                                           \n ";
            html += "         {                                                                           \n ";
            html += "             background: #6495ED;                                                    \n ";
            html += "         }                                                                           \n ";
            html += "         table td, table th                                                          \n ";
            html += "         {                                                                           \n ";
            html += "             border-right: 1px solid #000;                                           \n ";
            html += "             border-bottom: 1px solid #000;                                          \n ";
            html += "             padding: 3px;                                                           \n ";
            html += "             line-height: 1em;                                                       \n ";
            html += "             font-size: 0.7em;                                                       \n ";
            html += "             font-family: 'Lucida Grande' , Helvetica, Arial, Verdana, sans-serif;   \n ";
            html += "         }                                                                           \n ";
            html += "         table tr.odd th, table tr.odd td                                            \n ";
            html += "         {                                                                           \n ";
            html += "             background: #efefef;                                                    \n ";
            html += "         }                                                                           \n ";
            html += "         .font-Title                                                                  \n ";
            html += "         {                                                                           \n ";
            html += "             font-size: 0.8em;                                                       \n ";
            html += "             font-family: 'Lucida Grande' , Helvetica, Arial, Verdana, sans-serif;   \n ";
            html += "             color: #333;                                                            \n ";
            html += "         }                                                                           \n ";
            html += "         .auto-style1                                                                \n ";
            html += "         {                                                                           \n ";
            html += "            text-align: right;                                                       \n ";
            html += "         }                                                                           \n ";
            html += "         .auto-style2                                                                \n ";
            html += "         {                                                                           \n ";
            html += "             text-align: center;                                                     \n ";
            html += "             font-family: 'Lucida Grande' , Helvetica, Arial, Verdana, sans-serif;   \n ";
            html += "             font-weight:bold;                                                        \n ";
            html += "         }                                                                           \n ";
            html += "     </style>                                                                        \n ";
            html += " </head>                                                                             \n ";
            html += " <body>                                                                             \n ";
            html += "<h3>superstar!</h3>";
            html += "<div>";
            html += "<img src=cid:img width=\"1098\" height=\"693\"/>";
            html += "</div>";

            if (ds.Tables["DataSet5"].Rows.Count > 0)
            {

                html += "<div>  \n";
                html += "<div><h5>2)고객사별/제품군별 매출현황(단위:백만원)</h5></div>  \n";
                html += "         <table>                                                                     \n";
                html += "    <thead>                                                                          \n";
                html += "                     <tr>                                                            \n";
                html += "                         <th scope='col' width='230' height='25'>                     \n";
                html += "                             고객사/제품군                                           \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             DEVELOPER                                               \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             상품                                                    \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             BUMP STRIPPE                                            \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             COLOR DEVELO                                            \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             NANO STRIPPE                                            \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             ETCHANT                                                 \n";
                html += "                         </th>                                                       \n";
                html += "                         <th scope='col' width='100'>                                 \n";
                html += "                             합계                                                    \n";
                html += "                         </th>                                                       \n";
                html += "                     </tr>                                                           \n";
                html += "                 </thead>                                                            \n";
                html += "                 <tbody>                                                             \n";

                if (ds.Tables["DataSet5"].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables["DataSet5"].Rows.Count; i++)
                    {
                        html += "                     <tr>                                                                          \n";
                        html += "                         <td class=\"auto-style2\">" + ds.Tables["DataSet5"].Rows[i]["BP_NM"].ToString() + "</td>        \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["DEVELOPER"].ToString() + "</td>    \n";
                        html += "                         <td class=\"auto-style1\"> " + ds.Tables["DataSet5"].Rows[i]["상품"].ToString() + "</td>        \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["BUMP"].ToString() + "</td>         \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["COLOR"].ToString() + "</td>        \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["NANO"].ToString() + "</td>         \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["ETCHANT"].ToString() + "</td>      \n";
                        html += "                         <td class=\"auto-style1\">" + ds.Tables["DataSet5"].Rows[i]["합계"].ToString() + "</td>         \n";
                        html += "                     </tr>                                                                         \n";
                    }
                }
                html += "                 </tbody>                                                            \n";
                html += "         </table>                                                                     \n";
                html += "</div>";
            }

            html += "<h5>이 메일은 발신전용 메일입니다. </h5>";
            html += "<h5>감사합니다. </h5>";
            html += "<h5>I'll serve you~</h5>";
            html += "</body></html>";

            AlternateView altView = AlternateView.CreateAlternateViewFromString(html, null, System.Net.Mime.MediaTypeNames.Text.Html);

            LinkedResource theEmailImage = new LinkedResource(Path, System.Net.Mime.MediaTypeNames.Image.Jpeg);
            theEmailImage.ContentId = "img";
            altView.LinkedResources.Add(theEmailImage);


            MailMessage msg = new MailMessage();  //매일 오전 10시에 자동실행
            msg.To.Add("ktjung@nepes.co.kr"); //정갑태
            //msg.To.Add("hphong@nepesamc.co.kr");  //홍학표            
            //msg.To.Add("stchung@nepes.co.kr");  //정성태      20160805 삭제요청 정지원      
            //msg.To.Add("davidkim@nepes.co.kr"); //김두희   20151218 삭제요청 박기수  
            //msg.To.Add("sm1092@nepes.co.kr"); //김상민 - 신제품 기획팀 20151218 삭제요청 박기수  
            msg.To.Add("leem@nepes.co.kr"); //이민  20170529 추가
            msg.To.Add("jhchoi@nepes.co.kr"); //최재훈
            msg.To.Add("red-fly@nepes.co.kr"); //문창규
            msg.To.Add("jjw7969@nepes.co.kr"); //정지원
            msg.To.Add("kimwd@nepes.co.kr"); //김원동

            /*문창규 차장 인원 추가요청 20151028*/
            //msg.To.Add("kjsx@nepes.co.kr");//em김정석  20160805 삭제요청 정지원
            msg.To.Add("pabian@nepes.co.kr");//em강지원
            //msg.To.Add("hurty@nepes.co.kr");//em허태영 20161209 삭제요청 문창규
            msg.To.Add("hongjy@nepes.co.kr");//em홍준영
            msg.To.Add("jhchoi@nepes.co.kr");//em최채훈
            //msg.To.Add("rhoshin@nepes.co.kr");//em신원조  20160805 삭제요청 김지일, 정지원
            msg.To.Add("elan6000@nepes.co.kr"); //em강승일
            msg.To.Add("hyuk7879@nepes.co.kr"); //em이종혁

            /*정지원 과장 인원 추가 요청 20160805*/
            msg.To.Add("jssong@nepes.co.kr"); //em송진석
            msg.To.Add("yjchoi0705@nepes.co.kr"); //em최유진           
            msg.To.Add("hsyoo@nepes.co.kr"); //em유홍석이사   20160805 추가요청 김지일

            /* 참조인원으로 변경요청 20151218 박기수*/
            msg.CC.Add("ywkim@nepes.co.kr"); // 김윤우
            msg.CC.Add("kspark@nepes.co.kr"); //박기수            
            msg.CC.Add("kimjg1001@nepes.co.kr"); //김준근 추가 20170529

            /*참조인원 추가요청 20151218 박기수*/
            msg.CC.Add("ktkim@nepes.co.kr"); // 김경태
            msg.CC.Add("tshyun@nepes.co.kr"); //현태수
            msg.CC.Add("leecw@nepes.co.kr"); //이창우
            msg.CC.Add("hklee@nepesamc.co.kr"); //이현규


            msg.Bcc.Add("janghs0501@nepes.co.kr");  // 장희성 bcc 숨은참조
            msg.Bcc.Add("ahncj@nepes.co.kr");  // 장희성 bcc 숨은참조
            msg.Bcc.Add("yoonst@nepes.co.kr");  // 윤승택 bcc 숨은참조

            //msg.CC.Add("yoosr@nepes.co.kr");   //cc는 참조
            msg.Subject = ReportViewer1.LocalReport.DisplayName;
            msg.IsBodyHtml = true;
            msg.AlternateViews.Add(altView);


            /*report view 파일떨구기 for Excel*/
            Warning[] warnings2;
            string[] streamids2;

            string deviceInfo2 = string.Empty;
            string mimeType2 = string.Empty;
            string encoding2 = string.Empty;
            string extension2 = string.Empty;
            byte[] bytes2 = ReportViewer1.LocalReport.Render("Excel", deviceInfo2, out mimeType2, out encoding2, out extension2, out streamids2, out warnings2);

            string PathExcel = Server.MapPath(".") + @"\EM_매출현황.xls";
            FileStream fs2 = new FileStream(PathExcel, FileMode.Create, FileAccess.Write);
            fs2.Write(bytes2, 0, bytes2.Length);
            fs2.Close();

            /*excel 파일첨부*/
            Attachment excel = new Attachment(PathExcel);
            msg.Attachments.Add(excel);

            /*메일보내기*/
            var sClient = new SmtpClient("mail.nepes.co.kr", 25);
            //var sClient = new SmtpClient();
            sClient.UseDefaultCredentials = false;
            using (sClient as IDisposable)
            {
                sClient.Send(msg);
            }
            msg.Dispose();


            string close = @"<script type='text/javascript'>
                                            window.returnValue = true;
                                            window.open('','_self','');                                                            
                                            window.close();
                                            </script>";
            base.Response.Write(close);
        }
        

        protected void getData()
        {
            ReportViewer1.Reset();

            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();

                sql_cmd2 = sql_conn.CreateCommand();
                sql_cmd2.CommandType = CommandType.Text;
                sql_cmd2.CommandText = getSQL2();

                sql_cmd3 = sql_conn.CreateCommand();
                sql_cmd3.CommandType = CommandType.Text;
                sql_cmd3.CommandText = getSQL3();

                sql_cmd4 = sql_conn.CreateCommand();
                sql_cmd4.CommandType = CommandType.Text;
                sql_cmd4.CommandText = getSQL4();

                sql_cmd5 = sql_conn.CreateCommand();
                sql_cmd5.CommandType = CommandType.Text;
                sql_cmd5.CommandText = getSQL5();

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                    da.Fill(ds, "DataSet1");

                    SqlDataAdapter da2 = new SqlDataAdapter(sql_cmd2);
                    da2.Fill(ds, "DataSet2");

                    SqlDataAdapter da3 = new SqlDataAdapter(sql_cmd3);
                    da3.Fill(ds, "DataSet3");

                    SqlDataAdapter da4 = new SqlDataAdapter(sql_cmd4);
                    da4.Fill(ds, "DataSet4");

                    SqlDataAdapter da5 = new SqlDataAdapter(sql_cmd5);
                    da5.Fill(ds, "DataSet5");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open)
                        sql_conn.Close();
                }
                sql_conn.Close();

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s7001_sub.rdlc");
                ReportViewer1.LocalReport.DisplayName = "EM 매출현황_" + tb_yyyymm.Substring(0, 4) + "_" + tb_yyyymm.Substring(4, 2) + "_" + tb_yyyymm.Substring(6, 2);


                /*원그래프*/
                ReportDataSource rds2 = new ReportDataSource();
                rds2.Name = "DataSet2_1";
                DataTable dt2_1 = ds.Tables["DataSet2"].Copy();
                dt2_1.DefaultView.RowFilter = "RW = '1'";
                rds2.Value = dt2_1.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds2);

                ReportDataSource rds2_2 = new ReportDataSource();
                rds2_2.Name = "DataSet2_2";
                DataTable dt2_2 = ds.Tables["DataSet2"].Copy();
                dt2_2.DefaultView.RowFilter = "RW = '2'";
                rds2_2.Value = dt2_2.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds2_2);

                ReportDataSource rds2_3 = new ReportDataSource();
                rds2_3.Name = "DataSet2_3";
                DataTable dt2_3 = ds.Tables["DataSet2"].Copy();
                dt2_3.DefaultView.RowFilter = "RW = '3'";
                rds2_3.Value = dt2_3.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds2_3);                


                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                rds.Value = ds.Tables["DataSet1"];
                ReportViewer1.LocalReport.DataSources.Add(rds);


                ReportDataSource rds3 = new ReportDataSource();
                rds3.Name = "DataSet3";
                rds3.Value = ds.Tables["DataSet3"];
                ReportViewer1.LocalReport.DataSources.Add(rds3);

                ReportDataSource rds4 = new ReportDataSource();
                rds4.Name = "DataSet4";
                rds4.Value = ds.Tables["DataSet4"];
                ReportViewer1.LocalReport.DataSources.Add(rds4);   

            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
            }            

        }

        private string getSQL()
        {            
            string date = tb_yyyymm;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MAIL_MONTH  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL2()
        {            
            string date = tb_yyyymm;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MAIL_MONTH_GRAPH  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL3()
        {
            string date = tb_yyyymm;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MAIL_MONTH_DET  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL4()
        {
            string date = tb_yyyymm;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MAIL_MONTH_CHART  '" + date + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL5()
        {
            string date = tb_yyyymm;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_BP_MAIL_MONTH_GRID  '" + date + "'\n");
            return sbSQL.ToString();
        }
    }
}