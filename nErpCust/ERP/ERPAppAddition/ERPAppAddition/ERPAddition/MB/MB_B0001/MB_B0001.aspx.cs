using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Json;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ERPAppAddition.ERPAddition.MB.MB_B0001
{
    public partial class MB_B0001 : System.Web.UI.Page
    {
        string sql_cust_cd;
        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        int value;

        string date = "";
        string to_page = "1";
        string from_page = "1";

        protected void Page_Load(object sender, EventArgs e)
        {
            Server.ScriptTimeout = 600; // 기본값은 web.config 의 executionTimeout(초) 임

            /*1일 1회 실행 전일 데이터 생성  작업스케줄러로 아침 8시에 실행됨*/
            DateTime day = DateTime.Now.AddDays(-1);
           string fromDT = day.AddDays(-3).Year.ToString("0000") + "-" + day.AddDays(-3).Month.ToString("00") + "-" + day.AddDays(-3).Day.ToString("00");
           string toDT = day.Year.ToString("0000") + "-" + day.Month.ToString("00") + "-" + day.Day.ToString("00");
            

            ///*강제 데이터 생성을 위한 파라미터 set*/
           //string fromDT = "2017-08-01";
           //string toDT = "2017-08-15";
           
            for (DateTime frDT = Convert.ToDateTime(fromDT); frDT <= Convert.ToDateTime(toDT); frDT = frDT.AddDays(+1) ) 
            {
                /*마법노트 반환값 주소*/
                date = frDT.Year.ToString("0000") + frDT.Month.ToString("00") + frDT.Day.ToString("00");
               
               // DatabaseInput("https://stest-dot-mabup2.appspot.com/stat/group/query/1/" + date + "/" + date + "/b11128aded0c4464838df46401bcfc98");
               //DatabaseInput("https://mabup2.appspot.com/stat/group/query/1/" + date + "/" + date + "/b11128aded0c4464838df46401bcfc98");
                 DatabaseInput("https://mabup2.appspot.com/stat/group/query/1/" + date + "/" + date + "/" + from_page + "/" + to_page + "/b11128aded0c4464838df46401bcfc98");

                 
            }

            /*작업 정상처리시 화면 자동 닫기~*/
            string close = @"<script type='text/javascript'>
                                            window.returnValue = true;
                                            window.open('','_self','');                                                            
                                            window.close();
                                            </script>";
            base.Response.Write(close);
            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }


        private void DatabaseInput(string webURL)
        {   
            string ret = "";
            HttpWebResponse resp;
            Stream stream;
            StreamReader reader;

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(webURL);            
            req.UserAgent = HttpContext.Current.Request.ServerVariables["HTTP_USER_AGENT"].ToString();
            req.Credentials = CredentialCache.DefaultCredentials;

            //원하는 값이 나올때까지 무한루프 돌린다. 
            while (ret.Length < 100)
            {
                resp = (HttpWebResponse)req.GetResponse();
                stream = resp.GetResponseStream();
                reader = new StreamReader(stream);
                ret = reader.ReadToEnd();
                stream.Close();
                resp.Close();
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("name");            
            dt.Columns.Add("email");
            dt.Columns.Add("id");
            dt.Columns.Add("Date");
            dt.Columns.Add("작성글");
            dt.Columns.Add("나눔글");
            dt.Columns.Add("나눔된글");
            dt.Columns.Add("내댓글");

            if (ret.Length > 0)
            {
                JsonTextParser parser = new JsonTextParser();
                JObject obj = JObject.Parse(ret);

                try
                {
                    JArray array = JArray.Parse(obj["stats"].ToString());
                    foreach (JObject itemObj in array)
                    {
                        DataRow dr = dt.NewRow();
                        dr["name"] = itemObj["name"].ToString();
                        dr["email"] = itemObj["email"].ToString();

                        JArray array2 = JArray.Parse(itemObj["stats"].ToString());
                        if (array2.Count > 0)
                        {
                            foreach (JObject item2Obj in array2)
                            {
                                dr["id"] = item2Obj["id"].ToString();
                                dr["Date"] = item2Obj["baseDate"].ToString();
                                dr["작성글"] = item2Obj["ref01"].ToString();
                                dr["나눔글"] = item2Obj["ref02"].ToString();
                                dr["나눔된글"] = item2Obj["ref03"].ToString();
                                dr["내댓글"] = item2Obj["ref04"].ToString();
                            }
                        }
                        else
                        {
                            dr["id"] = " ";
                            dr["Date"] = date;
                        }
                        dt.Rows.Add(dr);
                    }
                }
                catch
                {
                    return;
                }
            }
            /*dt insert */
            if (dt.Rows.Count > 0)
            {
                StringBuilder sbSQL = new StringBuilder();
                /*해당일자 DATA 삭제후 */
                sbSQL.Append("DELETE  H_MABUPNOTE \n");
                sbSQL.Append("WHERE  YYYYMMDD = '" + date+ "'  \n");

                /*insert 쿼리를 rows count 만큼 한방에 만들어서 send*/
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sbSQL.Append("insert into H_MABUPNOTE \n");
                    sbSQL.Append("(                     \n");
                    sbSQL.Append(" YYYYMMDD             \n");
                    sbSQL.Append(",ID                   \n");
                    sbSQL.Append(",EMAIL                \n");
                    sbSQL.Append(",NAME                 \n");
                    sbSQL.Append(",MY_NOTE              \n");
                    sbSQL.Append(",SND_NOTE             \n");
                    sbSQL.Append(",RCV_NOTE             \n");
                    sbSQL.Append(",RE_NOTE              \n");
                    sbSQL.Append(",INSRT_DT             \n");
                    sbSQL.Append(",INSRT_USER           \n");
                    sbSQL.Append(",UPDT_DT              \n");
                    sbSQL.Append(",UPDT_USER            \n");
                    sbSQL.Append(")                     \n");
                    sbSQL.Append("VALUES(               \n");
                    sbSQL.Append(" '" + dt.Rows[i]["Date"] +     "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["ID"] + "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["email"] +    "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["name"] +     "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["작성글"] +   "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["나눔글"] +   "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["나눔된글"] + "' \n");
                    sbSQL.Append(",'" + dt.Rows[i]["내댓글"] +   "' \n");                    
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(",'MB_B0001'               \n");
                    sbSQL.Append(",GETDATE()             \n");
                    sbSQL.Append(",'MB_B0001'               \n");
                    sbSQL.Append(" )                     \n");
                }                 

                string setSQL = sbSQL.ToString();
                if (QueryExecute(setSQL) < 0)
                {
                    MessageBox.ShowMessage("입력 값에 오류 데이터가 있습니다.", this.Page);
                    return;
                }                
            }
        }


        public int QueryExecute(string sql)
        {
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandType = CommandType.Text;
            sql_cmd.CommandText = sql;

            try
            {
                value = sql_cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open)
                    sql_conn.Close();
                value = -1;
            }
            sql_conn.Close();
            return value;

        }
    }
}