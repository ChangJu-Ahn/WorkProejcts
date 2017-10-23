using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;

namespace ERPAppAddition.ERPAddition.B2.B_PSW
{
    public partial class B_PSW : System.Web.UI.Page
    {

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                setCombo();

            }
        }

        private void setCombo()
        {
            /*계열사선택*/
            DataTable dtFac = new DataTable();
            dtFac.Columns.Add("fac_cd", typeof(string));
            dtFac.Columns.Add("fac_nm", typeof(string));

            dtFac.Rows.Add("nepes", "네패스");
            //dtFac.Rows.Add("nepes_display", "Display");
            //dtFac.Rows.Add("nepes_led", "LED");
            //dtFac.Rows.Add("nepes_amc", "AMC");
            //dtFac.Rows.Add("nepes_enc", "ENC");
            dtFac.Rows.Add("nepes_test1", "TEST1");

            ddl_fac.DataTextField = "fac_nm";
            ddl_fac.DataValueField = "fac_cd";
            ddl_fac.DataSource = dtFac;
            ddl_fac.DataBind();
            ddl_fac.SelectedIndex = 0;

        }

        protected void send_Click(object sender, EventArgs e)
        {
            if (txt_id.Text.Equals("") || txt_mail.Text.Equals(""))
            {
                MessageBox.ShowMessage("id와 mail 을 입력하세요", this.Page);
                return;
            }

            DataSet ds = new DataSet();

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
                MessageBox.ShowMessage("해당 메일주소와 ID 정보가 없습니다. 입력하신 정보를 확인 바랍니다.", this.Page);
                return;
            }
            else
            {
                sql_cmd.CommandText = getSqlPw();

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
                //MessageBox.ShowMessage("사용자 등록 mail로 임시 비밀번호가 발급 되었습니다. Mail로 발송된 임시 비밀번호로 로그인 바랍니다. [Mail발송 대기시간 약 5분]", this.Page);
                ScriptManager.RegisterStartupScript(this.Page, this.GetType(), "Msg", "uPageClose();", true);
            }
        }

        private string getSQL()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            string dbName = ddl_fac.SelectedValue.ToString();
            string eMail = txt_mail.Text.ToString();
            string userId = txt_id.Text.ToString();
            
            sbSQL.Append("SELECT B.EMAIL_ADDR FROM " + dbName + ".dbo.Z_USR_MAST_REC A INNER JOIN " + dbName + ".dbo.HAA010T B ON A.USR_NM = B.NAME \n");
            sbSQL.Append("WHERE EMAIL_ADDR  = '" + eMail + "' AND USR_ID = '" + userId + "' \n");
            return sbSQL.ToString();
        }

        private string getSqlPw()
        {
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            string dbName = ddl_fac.SelectedValue.ToString();
            string eMail = txt_mail.Text.ToString();
            string userId = txt_id.Text.ToString();

            sbSQL.Append("USE "+dbName+" \n");
            sbSQL.Append("EXEC USP_B_PSW '" + dbName + "', '" + userId + "','" + eMail + "' \n");
            return sbSQL.ToString();
        }


        protected void exit_Click(object sender, EventArgs e)
        {

            string close = @"<script type='text/javascript'>
                                            window.returnValue = true;
                                            window.open('','_self','');                                                            
                                            window.close();
                                            </script>";
            base.Response.Write(close);
        }

    }
}