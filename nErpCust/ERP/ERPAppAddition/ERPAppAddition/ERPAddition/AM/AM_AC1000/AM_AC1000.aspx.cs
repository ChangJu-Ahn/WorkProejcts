using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.AM.AM_AC1000;

namespace ERPAppAddition.ERPAddition.AM.AM_AC1000
{
    public partial class AM_AC1000 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_display"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        string id = "";

        // 페이지 로드
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["id"] != null)
            {
                id = Request.QueryString["id"];
            }
            else
                id = "";

            WebSiteCount();
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        // 계산로직 (일 자금계획)
        protected void Cal_A_Value()
        {

            string LastMon = string.Empty;
            Double LastValue = 0;
            LastMon = "EXEC USP_A_DAILY_AMT_LASTMON '" + txt_yyyy.Text + "','" + txt_mm.Text + "', '" + txt_dd.Text + "'  ";
            cmd.Connection = conn;
            cmd.CommandText = LastMon;
            SqlDataReader dReader_LastMon = cmd.ExecuteReader();

            if (dReader_LastMon.HasRows)
            {
                dReader_LastMon.Read();
                LastValue = double.Parse(dReader_LastMon.GetValue(0).ToString());
            }

            else
            {
                LastValue = 0;
            }

            tb_amt8.Text = (Convert.ToDouble(tb_amt1.Text) + Convert.ToDouble(tb_amt2.Text) + Convert.ToDouble(tb_amt3.Text) + Convert.ToDouble(tb_amt4.Text) + Convert.ToDouble(tb_amt5.Text) + Convert.ToDouble(tb_amt6.Text) + Convert.ToDouble(tb_amt7.Text)).ToString(); // 영업활동상의 자금수입 합계
            tb_amt14.Text = (Convert.ToDouble(tb_amt10.Text) + Convert.ToDouble(tb_amt11.Text) + Convert.ToDouble(tb_amt12.Text) + Convert.ToDouble(tb_amt13.Text)).ToString(); // 인건비의 지급 합계
            tb_amt22.Text = (Convert.ToDouble(tb_amt14.Text) + Convert.ToDouble(tb_amt15.Text) + Convert.ToDouble(tb_amt16.Text) + Convert.ToDouble(tb_amt17.Text) + Convert.ToDouble(tb_amt18.Text) + Convert.ToDouble(tb_amt19.Text) + Convert.ToDouble(tb_amt20.Text) + Convert.ToDouble(tb_amt21.Text) + Convert.ToDouble(tb_amt9.Text)).ToString(); // 영업활동상의 자금지출
            tb_amt23.Text = (Convert.ToDouble(tb_amt8.Text) - Convert.ToDouble(tb_amt22.Text)).ToString(); //영업상의 순 자금유입 계산
            tb_amt28.Text = (Convert.ToDouble(tb_amt24.Text) + Convert.ToDouble(tb_amt25.Text) + Convert.ToDouble(tb_amt26.Text) + Convert.ToDouble(tb_amt27.Text)).ToString(); // 투자활동의 Cash in 합계
            tb_amt32.Text = (Convert.ToDouble(tb_amt29.Text) + Convert.ToDouble(tb_amt30.Text) + Convert.ToDouble(tb_amt31.Text)).ToString(); //고정자산의 취득
            tb_amt36.Text = (Convert.ToDouble(tb_amt32.Text) + Convert.ToDouble(tb_amt33.Text) + Convert.ToDouble(tb_amt34.Text) + Convert.ToDouble(tb_amt35.Text)).ToString(); // 투자활동의 Cash Out 계산
            tb_amt37.Text = (Convert.ToDouble(tb_amt28.Text) - Convert.ToDouble(tb_amt36.Text)).ToString(); // 투자활동의 자금흐름 계산
            tb_amt41.Text = (Convert.ToDouble(tb_amt38.Text) + Convert.ToDouble(tb_amt39.Text) + Convert.ToDouble(tb_amt40.Text)).ToString(); // 재무활동의 Cash in 합계
            tb_amt45.Text = (Convert.ToDouble(tb_amt42.Text) + Convert.ToDouble(tb_amt43.Text) + Convert.ToDouble(tb_amt44.Text)).ToString(); //재무활동의 Cash Out 합계
            tb_amt46.Text = (Convert.ToDouble(tb_amt41.Text) - Convert.ToDouble(tb_amt45.Text)).ToString(); // 재무활동의 현금흐름 계산
            tb_amt47.Text = (LastValue).ToString(); // 전월이월자금
            tb_amt48.Text = (Convert.ToDouble(tb_amt23.Text) + Convert.ToDouble(tb_amt37.Text) + Convert.ToDouble(tb_amt46.Text)).ToString();
            tb_amt49.Text = (Convert.ToDouble(tb_amt47.Text) + Convert.ToDouble(tb_amt48.Text)).ToString();

            dReader_LastMon.Close();
        }

        // 계산로직 (월 자금계획)
        protected void Cal_B_Value()
        {
            string LastMon_plan = string.Empty;
            Double LastPlanValue = 0;
            LastMon_plan = "EXEC USP_A_DAILY_AMT_PLAN_LASTMON '" + txt_yyyy.Text + "','" + txt_mm.Text + "', '01'  ";
            cmd.Connection = conn;
            cmd.CommandText = LastMon_plan;
            SqlDataReader dReader_LastMon_plan = cmd.ExecuteReader();

            if (dReader_LastMon_plan.HasRows)
            {
                dReader_LastMon_plan.Read();
                LastPlanValue = double.Parse(dReader_LastMon_plan.GetValue(0).ToString());
            }

            else
            {
                LastPlanValue = 0;
            }

            tb_amt_plan_3.Text = (Convert.ToDouble(tb_amt_plan_1.Text) + Convert.ToDouble(tb_amt_plan_2.Text)).ToString(); //영업현금유입
            tb_amt_plan_5.Text = (Convert.ToDouble(tb_amt_plan_6.Text) + Convert.ToDouble(tb_amt_plan_7.Text) + Convert.ToDouble(tb_amt_plan_8.Text) + Convert.ToDouble(tb_amt_plan_9.Text) + Convert.ToDouble(tb_amt_plan_10.Text) + Convert.ToDouble(tb_amt_plan_11.Text)).ToString(); // 미지급금
            tb_amt_plan_14.Text = (Convert.ToDouble(tb_amt_plan_4.Text) + Convert.ToDouble(tb_amt_plan_5.Text) + Convert.ToDouble(tb_amt_plan_12.Text) + Convert.ToDouble(tb_amt_plan_13.Text)).ToString(); // 영업현금유출
            tb_amt_plan_15.Text = (Convert.ToDouble(tb_amt_plan_3.Text) - Convert.ToDouble(tb_amt_plan_14.Text)).ToString(); // 영업현금수지
            tb_amt_plan_19.Text = (Convert.ToDouble(tb_amt_plan_16.Text) + Convert.ToDouble(tb_amt_plan_17.Text) + Convert.ToDouble(tb_amt_plan_18.Text)).ToString(); // 영업외현금유출
            tb_amt_plan_20.Text = (Convert.ToDouble(tb_amt_plan_15.Text) - Convert.ToDouble(tb_amt_plan_19.Text)).ToString(); // 당월총현금수지
            tb_amt_plan_21.Text = (LastPlanValue).ToString(); ;// 기초현금
            tb_amt_plan_22.Text = (Convert.ToDouble(tb_amt_plan_20.Text) + Convert.ToDouble(tb_amt_plan_21.Text)).ToString(); // 기말현금
            tb_amt_plan_23.Text = (Convert.ToDouble(tb_amt_plan_22.Text)).ToString(); //가용현금

            dReader_LastMon_plan.Close();
        }

        // 계산로직 (월 자금실적)
        protected void Cal_B_Value_RSLT()
        {
            string LastMon_plan_rslt = string.Empty;
            Double LastRSLTValue = 0;
            LastMon_plan_rslt = "EXEC USP_A_DAILY_AMT_PLAN_LASTMON_RSLT '" + txt_yyyy.Text + "','" + txt_mm.Text + "', '01'  ";
            cmd.Connection = conn;
            cmd.CommandText = LastMon_plan_rslt;
            SqlDataReader dReader_LastMon_plan_rslt = cmd.ExecuteReader();

            if (dReader_LastMon_plan_rslt.HasRows)
            {
                dReader_LastMon_plan_rslt.Read();
                LastRSLTValue = double.Parse(dReader_LastMon_plan_rslt.GetValue(0).ToString());
            }

            else
            {
                LastRSLTValue = 0;
            }

            tb_amt_plan_RSLT_3.Text = (Convert.ToDouble(tb_amt_plan_RSLT_1.Text) + Convert.ToDouble(tb_amt_plan_RSLT_2.Text)).ToString(); //영업현금유입
            tb_amt_plan_RSLT_5.Text = (Convert.ToDouble(tb_amt_plan_RSLT_6.Text) + Convert.ToDouble(tb_amt_plan_RSLT_7.Text) + Convert.ToDouble(tb_amt_plan_RSLT_8.Text) + Convert.ToDouble(tb_amt_plan_RSLT_9.Text) + Convert.ToDouble(tb_amt_plan_RSLT_10.Text) + Convert.ToDouble(tb_amt_plan_RSLT_11.Text)).ToString(); // 미지급금
            tb_amt_plan_RSLT_14.Text = (Convert.ToDouble(tb_amt_plan_RSLT_4.Text) + Convert.ToDouble(tb_amt_plan_RSLT_5.Text) + Convert.ToDouble(tb_amt_plan_RSLT_12.Text) + Convert.ToDouble(tb_amt_plan_RSLT_13.Text)).ToString(); // 영업현금유출
            tb_amt_plan_RSLT_15.Text = (Convert.ToDouble(tb_amt_plan_RSLT_3.Text) - Convert.ToDouble(tb_amt_plan_RSLT_14.Text)).ToString(); // 영업현금수지
            tb_amt_plan_RSLT_19.Text = (Convert.ToDouble(tb_amt_plan_RSLT_16.Text) + Convert.ToDouble(tb_amt_plan_RSLT_17.Text) + Convert.ToDouble(tb_amt_plan_RSLT_18.Text)).ToString(); // 영업외현금유출
            tb_amt_plan_RSLT_20.Text = (Convert.ToDouble(tb_amt_plan_RSLT_15.Text) - Convert.ToDouble(tb_amt_plan_RSLT_19.Text)).ToString(); // 당월총현금수지
            tb_amt_plan_RSLT_21.Text = (LastRSLTValue).ToString(); ;// 기초현금
            tb_amt_plan_RSLT_22.Text = (Convert.ToDouble(tb_amt_plan_RSLT_20.Text) + Convert.ToDouble(tb_amt_plan_RSLT_21.Text)).ToString(); // 기말현금
            tb_amt_plan_RSLT_23.Text = (Convert.ToDouble(tb_amt_plan_RSLT_22.Text)).ToString(); //가용현금

            dReader_LastMon_plan_rslt.Close();
        }

        // 화면단 메뉴숨김 함수
        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (rbl_view_type.SelectedValue)
            {
                case "A": // 일자금실적 등록
                    table.Visible = true;
                    ld_yyyy.Visible = true;
                    txt_yyyy.Visible = true;
                    ld_mm.Visible = true;
                    txt_mm.Visible = true;
                    ld_dd.Visible = true;
                    txt_dd.Visible = true;
                    td.Visible = true;
                    //ckb_hld.Visible = true;
                    Select_Button.Visible = true;
                    Save_Button.Visible = true;
                    Update_Button.Visible = true;
                    div_day_spread.Visible = true;
                    div_month_spread.Visible = false;
                    break;

                case "B": // 월자금실적 등록
                    table.Visible = true;
                    ld_yyyy.Visible = true;
                    txt_yyyy.Visible = true;
                    ld_mm.Visible = true;
                    txt_mm.Visible = true;
                    ld_dd.Visible = false;
                    txt_dd.Visible = false;
                    td.Visible = false;
                    //ckb_hld.Visible = true;
                    Select_Button.Visible = true;
                    Save_Button.Visible = true;
                    Update_Button.Visible = true;
                    div_day_spread.Visible = false;
                    div_month_spread.Visible = true;
                    break;
            }
        }

        // 조회버튼 로직
        protected void btn_Select_Click(object sender, EventArgs e)
        {
            string select_queryStr;
            string[] temp;

            if (rbl_view_type.SelectedValue == "A")
            {
                select_queryStr = "SELECT * FROM A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' and DD = '" + txt_dd.Text + "' ";
            }
            else
            {
                select_queryStr = "SELECT * FROM A_DAILY_AMT_PLAN WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' ";
            }

            conn.Open();

            switch (rbl_view_type.SelectedValue)
            {
                case "A":
                    if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 년 을 확인하세요')", true);
                        return;
                    }
                    if (txt_mm == null || txt_mm.Text.Equals(""))
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 월 을 확인하세요')", true);
                        return;
                    }
                    if (txt_dd == null || txt_dd.Text.Equals(""))
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 일 을 확인하세요')", true);
                        return;
                    }
                    break;

                case "B":
                    if (txt_yyyy == null || txt_yyyy.Text.Equals(""))
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 년 을 확인하세요')", true);
                        return;
                    }
                    if (txt_mm == null || txt_mm.Text.Equals(""))
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 월 을 확인하세요')", true);
                        return;
                    }
                    break;
            }

            cmd.Connection = conn;
            cmd.CommandText = select_queryStr;
            SqlDataReader dReader_select = cmd.ExecuteReader();

            temp = new string[dReader_select.FieldCount];

            if (dReader_select.Read())
            {
                for (int a = 0; a < dReader_select.FieldCount; a++)
                {
                    temp[a] = dReader_select[a].ToString();
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜의 데이터가 없습니다.')", true);
                return;
            }

            switch (rbl_view_type.SelectedValue)
            {
                case "A":

                    txt_yyyy.Text = temp[0];
                    txt_mm.Text = temp[1];
                    txt_dd.Text = temp[2];
                    tb_amt1.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[3]));
                    tb_amt2.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[4]));
                    tb_amt3.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[5]));
                    tb_amt4.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[6]));
                    tb_amt5.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[7]));
                    tb_amt6.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[8]));
                    tb_amt7.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[9]));
                    tb_amt8.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[10]));
                    tb_amt9.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[11]));
                    tb_amt10.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[12]));
                    tb_amt11.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[13]));
                    tb_amt12.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[14]));
                    tb_amt13.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[15]));
                    tb_amt14.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[16]));
                    tb_amt15.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[17]));
                    tb_amt16.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[18]));
                    tb_amt17.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[19]));
                    tb_amt18.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[20]));
                    tb_amt19.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[21]));
                    tb_amt20.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[22]));
                    tb_amt21.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[23]));
                    tb_amt22.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[24]));
                    tb_amt23.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[25]));
                    tb_amt24.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[26]));
                    tb_amt25.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[27]));
                    tb_amt26.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[28]));
                    tb_amt27.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[29]));
                    tb_amt28.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[30]));
                    tb_amt29.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[31]));
                    tb_amt30.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[32]));
                    tb_amt31.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[33]));
                    tb_amt32.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[34]));
                    tb_amt33.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[35]));
                    tb_amt34.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[36]));
                    tb_amt35.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[37]));
                    tb_amt36.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[38]));
                    tb_amt37.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[39]));
                    tb_amt38.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[40]));
                    tb_amt39.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[41]));
                    tb_amt40.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[42]));
                    tb_amt41.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[43]));
                    tb_amt42.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[44]));
                    tb_amt43.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[45]));
                    tb_amt44.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[46]));
                    tb_amt45.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[47]));
                    tb_amt46.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[48]));
                    tb_amt47.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[49]));
                    tb_amt48.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[50]));
                    tb_amt49.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[51]));

                    break;

                case "B":
                    txt_yyyy.Text = temp[0];
                    txt_mm.Text = temp[1];
                    tb_amt_plan_1.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[3]));
                    tb_amt_plan_2.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[5]));
                    tb_amt_plan_3.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[7]));
                    tb_amt_plan_4.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[9]));
                    tb_amt_plan_5.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[11]));
                    tb_amt_plan_6.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[13]));
                    tb_amt_plan_7.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[15]));
                    tb_amt_plan_8.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[17]));
                    tb_amt_plan_9.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[19]));
                    tb_amt_plan_10.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[21]));
                    tb_amt_plan_11.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[23]));
                    tb_amt_plan_12.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[25]));
                    tb_amt_plan_13.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[27]));
                    tb_amt_plan_14.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[29]));
                    tb_amt_plan_15.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[31]));
                    tb_amt_plan_16.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[33]));
                    tb_amt_plan_17.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[35]));
                    tb_amt_plan_18.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[37]));
                    tb_amt_plan_19.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[39]));
                    tb_amt_plan_20.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[41]));
                    tb_amt_plan_21.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[43]));
                    tb_amt_plan_22.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[45]));
                    tb_amt_plan_23.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[47]));
                    tb_amt_plan_RSLT_1.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[4]));
                    tb_amt_plan_RSLT_2.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[6]));
                    tb_amt_plan_RSLT_3.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[8]));
                    tb_amt_plan_RSLT_4.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[10]));
                    tb_amt_plan_RSLT_5.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[12]));
                    tb_amt_plan_RSLT_6.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[14]));
                    tb_amt_plan_RSLT_7.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[16]));
                    tb_amt_plan_RSLT_8.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[18]));
                    tb_amt_plan_RSLT_9.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[20]));
                    tb_amt_plan_RSLT_10.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[22]));
                    tb_amt_plan_RSLT_11.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[24]));
                    tb_amt_plan_RSLT_12.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[26]));
                    tb_amt_plan_RSLT_13.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[28]));
                    tb_amt_plan_RSLT_14.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[30]));
                    tb_amt_plan_RSLT_15.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[32]));
                    tb_amt_plan_RSLT_16.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[34]));
                    tb_amt_plan_RSLT_17.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[36]));
                    tb_amt_plan_RSLT_18.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[38]));
                    tb_amt_plan_RSLT_19.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[40]));
                    tb_amt_plan_RSLT_20.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[42]));
                    tb_amt_plan_RSLT_21.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[44]));
                    tb_amt_plan_RSLT_22.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[46]));
                    tb_amt_plan_RSLT_23.Text = string.Format("{0:#,0}", Convert.ToDouble(temp[48]));

                    break;
            }

            Save_Button.Enabled = false;

            conn.Close();
            dReader_select.Close();
            conn.Dispose();

        }

        // 저장버튼 로직
        protected void btn_Save_Click(object sender, EventArgs e)
        {
            switch (rbl_view_type.SelectedValue)
            {
                case "A":
                    if (txt_yyyy.Text == "" ||
                        txt_mm.Text == "" ||
                        txt_dd.Text == ""
                       )
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('년, 월, 일을 확인하세요.')", true);
                        return;
                    }
                    break;

                case "B":
                    if (txt_yyyy.Text == "" ||
                        txt_mm.Text == ""
                        )
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('년, 월을 확인하세요.')", true);
                        return;
                    }
                    break;
            }

            conn.Open();
            cmd.Connection = conn;
            SqlTransaction tran_insert = conn.BeginTransaction();
            cmd.Transaction = tran_insert;

            string select_data = "";

            if (rbl_view_type.SelectedValue == "A")
            {
                select_data = " SELECT * FROM A_DAILY_AMT WHERE YYYY+MM+DD >= '" + txt_yyyy.Text + "" + txt_mm.Text + "" + txt_dd.Text + "'";
            }
            else
            {
                select_data = " SELECT * FROM A_DAILY_AMT_PLAN WHERE YYYY+MM >= '" + txt_yyyy.Text + "" + txt_mm.Text + "'";
            }

            cmd.CommandText = select_data;
            SqlDataReader dt_select = cmd.ExecuteReader();

            if (dt_select.Read())
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당일자에 데이터가 존재하거나, 입력일 이후 데이터가 존재합니다.')", true);
                dt_select.Close(); // 데이터가 있는지 확인을 위한 리더를 닫는 문 (데이터가 있을경우 팝업창 후 닫는 문)
                return;
            }

            dt_select.Close(); // 데이터가 있는지 확인을 위한 리더를 닫는 문 (데이터가 없을경우 닫는 문) 

            try
            {
                string insert_queryStr;

                if (rbl_view_type.SelectedValue == "A")
                {
                    Cal_A_Value(); // 콤보박스에 따라 다른 계산식 수행 (일 자금실적 계산)

                    insert_queryStr =
                        @"INSERT INTO [nepes_display].[dbo].[A_DAILY_AMT]
                                           ([YYYY]
                                           ,[MM]
                                           ,[DD]
                                           ,[LENDER_COLLET]
                                           ,[SURTAX_REFUND]
                                           ,[TARIFF_REFUND]
                                           ,[LEASE_INCOME]
                                           ,[IMPORT_INCOME]
                                           ,[USELESS_PAY]
                                           ,[ETC_1]
                                           ,[ETC_1_BUSINESS_INCOME]
                                           ,[M_MATERIAL]
                                           ,[M_PAY]
                                           ,[M_RETIRE]
                                           ,[M_FOUNTAIN]
                                           ,[M_WELRARE]
                                           ,[PAYROLL_COSTS_EXPENSE]
                                           ,[EXPENSE]
                                           ,[OS_PAY]
                                           ,[OUTSIDE_PAY]
                                           ,[EXPENSE_ETC]
                                           ,[INSEREST]
                                           ,[DEPOSIT_LEASE]
                                           ,[ETC_2]
                                           ,[ETC_2_BUSINESS_EXPENSE]
                                           ,[ETC_2_BUSINESS_INFLOW]
                                           ,[FIXED_ASSET_OUT]
                                           ,[LOAN_INCOME]
                                           ,[FINANCIAL_IN]
                                           ,[VALUABLE_PAPAERS_INCOME]
                                           ,[INVERST]
                                           ,[INVEST_LAND]
                                           ,[INVEST_MACHINE]
                                           ,[INVEST_CAR]
                                           ,[FIXED_ASSET_IN]
                                           ,[LOAN_EXPENSE]
                                           ,[FINANCIAL_OUT]
                                           ,[INVERST_OUT]
                                           ,[INVEST_OUT_CASHOUT]
                                           ,[INVEST_OUT_CASH]
                                           ,[BEWBORROW]
                                           ,[USANCE_CD1]
                                           ,[INCRES_CAPITAL]
                                           ,[INCRES_CAPITAL_CASH]
                                           ,[LOAN_CD1]
                                           ,[USANCE_CD2]
                                           ,[DIVIDEND]
                                           ,[DIVIDEND_CASHOUT]
                                           ,[DIVIDEND_CASH]
                                           ,[DIVIDEND_LAST_AMT]
                                           ,[DIVIDEND_AMT]
                                           ,[DIVIDEND_OVER_AMT]
                                           ,[REMARK]
                                           ,[INSRT_USER_ID]
                                           ,[INSRT_DT]
                                           ,[UPDT_USER_ID]
                                           ,[UPDT_DT])
                                     VALUES (";
                    insert_queryStr += "'" + txt_yyyy.Text + "',";
                    insert_queryStr += "'" + txt_mm.Text + "',";
                    insert_queryStr += "'" + txt_dd.Text + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt1.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt2.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt3.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt4.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt5.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt6.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt7.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt8.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt9.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt10.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt11.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt12.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt13.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt14.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt15.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt16.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt17.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt18.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt19.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt20.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt21.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt22.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt23.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt24.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt25.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt26.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt27.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt28.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt29.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt30.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt31.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt32.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt33.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt34.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt35.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt36.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt37.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt38.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt39.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt40.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt41.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt42.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt43.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt44.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt45.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt46.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt47.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt48.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt49.Text) + "',";
                    insert_queryStr += "'',";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate(),";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate())";
                }

                else
                {
                    Cal_B_Value(); // 콤보박스에 따라 다른 계산식 수행 (월 자금계획 값 계산)
                    Cal_B_Value_RSLT(); // (월 자금계획 실적 값 계산)

                    insert_queryStr =
                        @"INSERT INTO [nepes_display].[dbo].[A_DAILY_AMT_PLAN]
                                           ([YYYY]
                                           ,[MM]
                                           ,[DD]
                                           ,[SALE_CREDIT]
                                           ,[RETURN_PAY]
                                           ,[SALES_PAY_IN]
                                           ,[TRADE_PAYABLE]
                                           ,[NONPAYMENT]
                                           ,[NONPAYMENT_OS]
                                           ,[NONPAYMENT_OUTSIDE]
                                           ,[NONPAYMENT_ETC1]
                                           ,[NONPAYMENT_ETC2]
                                           ,[NONPAYMENT_MATERIAL]
                                           ,[NONPAYMENT_ETCPAY]
                                           ,[PERSONNEL_EXPENSES]
                                           ,[PAID_INTEREST]
                                           ,[BUSINESS_CASH_OUT]
                                           ,[BUSINESS_CASH_FLOW]
                                           ,[FIXED_ASSETS]
                                           ,[ETC_BUSINESS_OUT]
                                           ,[ETC_BUSINESS_IN]
                                           ,[SALES_PAY_ETC]
                                           ,[SPOTMONTH_PAY]
                                           ,[BASIS_PAY]
                                           ,[TERMEND_PAY]
                                           ,[AVAILABLE_PAY]
                                           ,[REMARK]
                                           ,[INSRT_USER_ID]
                                           ,[INSRT_DT]
                                           ,[UPDT_USER_ID]
                                           ,[UPDT_DT]
                                           ,[SALE_CREDIT_RSLT]
                                           ,[RETURN_PAY_RSLT]
                                           ,[SALES_PAY_IN_RSLT]
                                           ,[TRADE_PAYABLE_RSLT]
                                           ,[NONPAYMENT_RSLT]
                                           ,[NONPAYMENT_OS_RSLT]
                                           ,[NONPAYMENT_OUTSIDE_RSLT]
                                           ,[NONPAYMENT_ETC1_RSLT]
                                           ,[NONPAYMENT_ETC2_RSLT]
                                           ,[NONPAYMENT_MATERIAL_RSLT]
                                           ,[NONPAYMENT_ETCPAY_RSLT]
                                           ,[PERSONNEL_EXPENSES_RSLT]
                                           ,[PAID_INTEREST_RSLT]
                                           ,[BUSINESS_CASH_OUT_RSLT]
                                           ,[BUSINESS_CASH_FLOW_RSLT]
                                           ,[FIXED_ASSETS_RSLT]
                                           ,[ETC_BUSINESS_OUT_RSLT]
                                           ,[ETC_BUSINESS_IN_RSLT]
                                           ,[SALES_PAY_ETC_RSLT]
                                           ,[SPOTMONTH_PAY_RSLT]
                                           ,[BASIS_PAY_RSLT]
                                           ,[TERMEND_PAY_RSLT]
                                           ,[AVAILABLE_PAY_RSLT])
                                     VALUES (";
                    insert_queryStr += "'" + txt_yyyy.Text + "',";
                    insert_queryStr += "'" + txt_mm.Text + "',";
                    insert_queryStr += "'01',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_1.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_2.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_3.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_4.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_5.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_6.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_7.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_8.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_9.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_10.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_11.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_12.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_13.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_14.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_15.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_16.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_17.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_18.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_19.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_20.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_21.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_22.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_23.Text) + "',";
                    insert_queryStr += "'',";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate(),";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate(),";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_1.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_2.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_3.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_4.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_5.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_6.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_7.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_8.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_9.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_10.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_11.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_12.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_13.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_14.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_15.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_16.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_17.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_18.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_19.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_20.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_21.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_22.Text) + "',";
                    insert_queryStr += "'" + Convert.ToDouble(tb_amt_plan_RSLT_23.Text) + "')";
                }

                cmd.CommandText = insert_queryStr;
                cmd.ExecuteNonQuery();
                tran_insert.Commit();
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('저장되었습니다.')", true);

            }

            catch
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('해당 날짜에 저장된 데이터가 존재합니다.')", true);
                tran_insert.Rollback();
            }

            conn.Close();
        }

        // 수정버튼 로직
        protected void btn_Update_Click(object sender, EventArgs e)
        {
            conn.Open();
            string updt_query = "";
            SqlTransaction tran_update = conn.BeginTransaction();
            cmd.Connection = conn;
            cmd.Transaction = tran_update;
            int Rowcnt;

            switch (rbl_view_type.SelectedValue)
            {
                case "A":
                    if (txt_yyyy.Text == "" ||
                        txt_mm.Text == "" ||
                        txt_dd.Text == "" ||
                        tb_amt1.Text == "" ||
                        tb_amt2.Text == "" ||
                        tb_amt3.Text == "" ||
                        tb_amt4.Text == "" ||
                        tb_amt5.Text == "" ||
                        tb_amt6.Text == "" ||
                        tb_amt7.Text == "" ||
                        tb_amt9.Text == "" ||
                        tb_amt10.Text == "" ||
                        tb_amt11.Text == "" ||
                        tb_amt12.Text == "" ||
                        tb_amt13.Text == "" ||
                        tb_amt15.Text == "" ||
                        tb_amt16.Text == "" ||
                        tb_amt17.Text == "" ||
                        tb_amt18.Text == "" ||
                        tb_amt19.Text == "" ||
                        tb_amt20.Text == "" ||
                        tb_amt21.Text == "" ||
                        tb_amt24.Text == "" ||
                        tb_amt25.Text == "" ||
                        tb_amt26.Text == "" ||
                        tb_amt27.Text == "" ||
                        tb_amt29.Text == "" ||
                        tb_amt30.Text == "" ||
                        tb_amt31.Text == "" ||
                        tb_amt33.Text == "" ||
                        tb_amt34.Text == "" ||
                        tb_amt35.Text == "" ||
                        tb_amt38.Text == "" ||
                        tb_amt39.Text == "" ||
                        tb_amt40.Text == "" ||
                        tb_amt42.Text == "" ||
                        tb_amt43.Text == "" ||
                        tb_amt44.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 일 을 확인하세요')", true);
                        return;
                    }

                    Cal_A_Value(); // 계산로직 함수 실행

                    updt_query = " UPDATE [nepes_display].[dbo].[A_DAILY_AMT] SET ";

                    updt_query += " [LENDER_COLLET] = '" + Convert.ToDouble(tb_amt1.Text) + "',";
                    updt_query += " [SURTAX_REFUND] = '" + Convert.ToDouble(tb_amt2.Text) + "',";
                    updt_query += " [TARIFF_REFUND] = '" + Convert.ToDouble(tb_amt3.Text) + "',";
                    updt_query += " [LEASE_INCOME] = '" + Convert.ToDouble(tb_amt4.Text) + "',";
                    updt_query += " [IMPORT_INCOME] = '" + Convert.ToDouble(tb_amt5.Text) + "',";
                    updt_query += " [USELESS_PAY] = '" + Convert.ToDouble(tb_amt6.Text) + "',";
                    updt_query += " [ETC_1] = '" + Convert.ToDouble(tb_amt7.Text) + "',";
                    updt_query += " [ETC_1_BUSINESS_INCOME] = '" + Convert.ToDouble(tb_amt8.Text) + "',";
                    updt_query += " [M_MATERIAL] = '" + Convert.ToDouble(tb_amt9.Text) + "',";
                    updt_query += " [M_PAY] = '" + Convert.ToDouble(tb_amt10.Text) + "',";
                    updt_query += " [M_RETIRE] = '" + Convert.ToDouble(tb_amt11.Text) + "',";
                    updt_query += " [M_FOUNTAIN] = '" + Convert.ToDouble(tb_amt12.Text) + "',";
                    updt_query += " [M_WELRARE] = '" + Convert.ToDouble(tb_amt13.Text) + "',";
                    updt_query += " [PAYROLL_COSTS_EXPENSE] = '" + Convert.ToDouble(tb_amt14.Text) + "',";
                    updt_query += " [EXPENSE] = '" + Convert.ToDouble(tb_amt15.Text) + "',";
                    updt_query += " [OS_PAY] = '" + Convert.ToDouble(tb_amt16.Text) + "',";
                    updt_query += " [OUTSIDE_PAY] = '" + Convert.ToDouble(tb_amt17.Text) + "',";
                    updt_query += " [EXPENSE_ETC] = '" + Convert.ToDouble(tb_amt18.Text) + "',";
                    updt_query += " [INSEREST] = '" + Convert.ToDouble(tb_amt19.Text) + "',";
                    updt_query += " [DEPOSIT_LEASE] = '" + Convert.ToDouble(tb_amt20.Text) + "',";
                    updt_query += " [ETC_2] = '" + Convert.ToDouble(tb_amt21.Text) + "',";
                    updt_query += " [ETC_2_BUSINESS_EXPENSE] = '" + Convert.ToDouble(tb_amt22.Text) + "',";
                    updt_query += " [ETC_2_BUSINESS_INFLOW] = '" + Convert.ToDouble(tb_amt23.Text) + "',";
                    updt_query += " [FIXED_ASSET_OUT] = '" + Convert.ToDouble(tb_amt24.Text) + "',";
                    updt_query += " [LOAN_INCOME] = '" + Convert.ToDouble(tb_amt25.Text) + "',";
                    updt_query += " [FINANCIAL_IN] = '" + Convert.ToDouble(tb_amt26.Text) + "',";
                    updt_query += " [VALUABLE_PAPAERS_INCOME] = '" + Convert.ToDouble(tb_amt27.Text) + "',";
                    updt_query += " [INVERST] = '" + Convert.ToDouble(tb_amt28.Text) + "',";
                    updt_query += " [INVEST_LAND] = '" + Convert.ToDouble(tb_amt29.Text) + "',";
                    updt_query += " [INVEST_MACHINE] = '" + Convert.ToDouble(tb_amt30.Text) + "',";
                    updt_query += " [INVEST_CAR] = '" + Convert.ToDouble(tb_amt31.Text) + "',";
                    updt_query += " [FIXED_ASSET_IN] = '" + Convert.ToDouble(tb_amt32.Text) + "',";
                    updt_query += " [LOAN_EXPENSE] = '" + Convert.ToDouble(tb_amt33.Text) + "',";
                    updt_query += " [FINANCIAL_OUT] = '" + Convert.ToDouble(tb_amt34.Text) + "',";
                    updt_query += " [INVERST_OUT] = '" + Convert.ToDouble(tb_amt35.Text) + "',";
                    updt_query += " [INVEST_OUT_CASHOUT] = '" + Convert.ToDouble(tb_amt36.Text) + "',";
                    updt_query += " [INVEST_OUT_CASH] = '" + Convert.ToDouble(tb_amt37.Text) + "',";
                    updt_query += " [BEWBORROW] = '"  + Convert.ToDouble(tb_amt38.Text) + "',";
                    updt_query += " [USANCE_CD1] = '" + Convert.ToDouble(tb_amt39.Text) + "',";
                    updt_query += " [INCRES_CAPITAL] = '" + Convert.ToDouble(tb_amt40.Text) + "',";
                    updt_query += " [INCRES_CAPITAL_CASH] = '" + Convert.ToDouble(tb_amt41.Text) + "',";
                    updt_query += " [LOAN_CD1] = '" + Convert.ToDouble(tb_amt42.Text) + "',";
                    updt_query += " [USANCE_CD2] = '" + Convert.ToDouble(tb_amt43.Text) + "',";
                    updt_query += " [DIVIDEND] = '" + Convert.ToDouble(tb_amt44.Text) + "',";
                    updt_query += " [DIVIDEND_CASHOUT] = '" + Convert.ToDouble(tb_amt45.Text) + "',";
                    updt_query += " [DIVIDEND_CASH] = '" + Convert.ToDouble(tb_amt46.Text) + "',";
                    updt_query += " [DIVIDEND_LAST_AMT] = '" + Convert.ToDouble(tb_amt47.Text) + "',";
                    updt_query += " [DIVIDEND_AMT] = '" + Convert.ToDouble(tb_amt48.Text) + "',";
                    updt_query += " [DIVIDEND_OVER_AMT] = '" + Convert.ToDouble(tb_amt49.Text) + "',";
                    updt_query += " [UPDT_USER_ID] = '" + id + "',";
                    updt_query += " [UPDT_DT] = getdate()";
                    updt_query += " WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' and DD = '" + txt_dd.Text + "'   ";

                    break;

                case "B":
                    if (txt_yyyy.Text == "" ||
                        txt_mm.Text == "" ||
                        tb_amt_plan_1.Text == "" ||
                        tb_amt_plan_2.Text == "" ||
                        tb_amt_plan_4.Text == "" ||
                        tb_amt_plan_6.Text == "" ||
                        tb_amt_plan_7.Text == "" ||
                        tb_amt_plan_8.Text == "" ||
                        tb_amt_plan_9.Text == "" ||
                        tb_amt_plan_10.Text == "" ||
                        tb_amt_plan_11.Text == "" ||
                        tb_amt_plan_12.Text == "" ||
                        tb_amt_plan_13.Text == "" ||
                        tb_amt_plan_16.Text == "" ||
                        tb_amt_plan_17.Text == "" ||
                        tb_amt_plan_18.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력 일 을 확인하세요')", true);
                        return;
                    }

                    Cal_B_Value();
                    Cal_B_Value_RSLT();

                    updt_query = " UPDATE [nepes_display].[dbo].[A_DAILY_AMT_PLAN] SET ";
                    updt_query += " [SALE_CREDIT] = '" + Convert.ToDouble(tb_amt_plan_1.Text) + "',";
                    updt_query += " [SALE_CREDIT_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_1.Text) + "',";
                    updt_query += " [RETURN_PAY] = '" + Convert.ToDouble(tb_amt_plan_2.Text) + "',";
                    updt_query += " [RETURN_PAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_2.Text) + "',";
                    updt_query += " [SALES_PAY_IN] = '" + Convert.ToDouble(tb_amt_plan_3.Text) + "',";
                    updt_query += " [SALES_PAY_IN_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_3.Text) + "',";
                    updt_query += " [TRADE_PAYABLE] = '" + Convert.ToDouble(tb_amt_plan_4.Text) + "',";
                    updt_query += " [TRADE_PAYABLE_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_4.Text) + "',";
                    updt_query += " [NONPAYMENT] = '" + Convert.ToDouble(tb_amt_plan_5.Text) + "',";
                    updt_query += " [NONPAYMENT_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_5.Text) + "',";
                    updt_query += " [NONPAYMENT_OS] = '" + Convert.ToDouble(tb_amt_plan_6.Text) + "',";
                    updt_query += " [NONPAYMENT_OS_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_6.Text) + "',";
                    updt_query += " [NONPAYMENT_OUTSIDE] = '" + Convert.ToDouble(tb_amt_plan_7.Text) + "',";
                    updt_query += " [NONPAYMENT_OUTSIDE_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_7.Text) + "',";
                    updt_query += " [NONPAYMENT_ETC1] = '" + Convert.ToDouble(tb_amt_plan_8.Text) + "',";
                    updt_query += " [NONPAYMENT_ETC1_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_8.Text) + "',";
                    updt_query += " [NONPAYMENT_ETC2] = '" + Convert.ToDouble(tb_amt_plan_9.Text) + "',";
                    updt_query += " [NONPAYMENT_ETC2_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_9.Text) + "',";
                    updt_query += " [NONPAYMENT_MATERIAL] = '" + Convert.ToDouble(tb_amt_plan_10.Text) + "',";
                    updt_query += " [NONPAYMENT_MATERIAL_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_10.Text) + "',";
                    updt_query += " [NONPAYMENT_ETCPAY] = '" + Convert.ToDouble(tb_amt_plan_11.Text) + "',";
                    updt_query += " [NONPAYMENT_ETCPAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_11.Text) + "',";
                    updt_query += " [PERSONNEL_EXPENSES] = '" + Convert.ToDouble(tb_amt_plan_12.Text) + "',";
                    updt_query += " [PERSONNEL_EXPENSES_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_12.Text) + "',";
                    updt_query += " [PAID_INTEREST] = '" + Convert.ToDouble(tb_amt_plan_13.Text) + "',";
                    updt_query += " [PAID_INTEREST_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_13.Text) + "',";
                    updt_query += " [BUSINESS_CASH_OUT] = '" + Convert.ToDouble(tb_amt_plan_14.Text) + "',";
                    updt_query += " [BUSINESS_CASH_OUT_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_14.Text) + "',";
                    updt_query += " [BUSINESS_CASH_FLOW] = '" + Convert.ToDouble(tb_amt_plan_15.Text) + "',";
                    updt_query += " [BUSINESS_CASH_FLOW_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_15.Text) + "',";
                    updt_query += " [FIXED_ASSETS] = '" + Convert.ToDouble(tb_amt_plan_16.Text) + "',";
                    updt_query += " [FIXED_ASSETS_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_16.Text) + "',";
                    updt_query += " [ETC_BUSINESS_OUT] = '" + Convert.ToDouble(tb_amt_plan_17.Text) + "',";
                    updt_query += " [ETC_BUSINESS_OUT_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_17.Text) + "',";
                    updt_query += " [ETC_BUSINESS_IN] = '" + Convert.ToDouble(tb_amt_plan_18.Text) + "',";
                    updt_query += " [ETC_BUSINESS_IN_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_18.Text) + "',";
                    updt_query += " [SALES_PAY_ETC] = '" + Convert.ToDouble(tb_amt_plan_19.Text) + "',";
                    updt_query += " [SALES_PAY_ETC_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_19.Text) + "',";
                    updt_query += " [SPOTMONTH_PAY] = '" + Convert.ToDouble(tb_amt_plan_20.Text) + "',";
                    updt_query += " [SPOTMONTH_PAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_20.Text) + "',";
                    updt_query += " [BASIS_PAY] = '" + Convert.ToDouble(tb_amt_plan_21.Text) + "',";
                    updt_query += " [BASIS_PAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_21.Text) + "',";
                    updt_query += " [TERMEND_PAY] = '" + Convert.ToDouble(tb_amt_plan_22.Text) + "',";
                    updt_query += " [TERMEND_PAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_22.Text) + "',";
                    updt_query += " [AVAILABLE_PAY] = '" + Convert.ToDouble(tb_amt_plan_23.Text) + "',";
                    updt_query += " [AVAILABLE_PAY_RSLT] = '" + Convert.ToDouble(tb_amt_plan_RSLT_23.Text) + "',";
                    updt_query += " [UPDT_USER_ID] = '" + id + "',";
                    updt_query += " [UPDT_DT] = getdate()";
                    updt_query += " WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' and DD = '01'   ";
                    
                    break;
            }

            try
            {
                cmd.CommandText = updt_query;
                Rowcnt = cmd.ExecuteNonQuery();

                if (Rowcnt == 0)
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('수정할 데이터가 없습니다.')", true);
                }

                else
                {
                    tran_update.Commit();
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('수정되었습니다.')", true);
                }  
            }

            catch
            {
                tran_update.Rollback();
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력되지 않은 항목이 있습니다.')", true);
            }
            
            if (rbl_view_type.SelectedValue == "A")
            {
                updt_trg();
            }
            else
            {
                updt_trg_plan_value();
                updt_trg_plan_value_RSLT();
            }
            conn.Close();

        }

        // 수정시 트리거 로직 (일자금 로직)
        protected void updt_trg()
        {
            string Today_AMT = "";
            DataSet1TableAdapters.USP_A_DAILY_AMT_NEXTDAYTableAdapter adapter = new DataSet1TableAdapters.USP_A_DAILY_AMT_NEXTDAYTableAdapter();
            DataSet1.USP_A_DAILY_AMT_NEXTDAYDataTable dt = adapter.GetData(txt_yyyy.Text, txt_mm.Text, txt_dd.Text);
            // USP_A_DAILY_AMT_NEXTDAY 프로시저를 사용하여 해당 수정일 이후의 모든 데이터를 끌고온다 (전월, 금월, 차월 금액)

            Today_AMT = " SELECT DIVIDEND_OVER_AMT FROM A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "'AND DD = '" + txt_dd.Text + "' ";
            // 수정된 차월데이터를 끌고와서 다음일 전월 데이터로 사용

            cmd.Connection = conn;
            cmd.CommandText = Today_AMT;
            SqlDataReader dReader_Today_AMT = cmd.ExecuteReader();

            double strVa = 0;

            if (dReader_Today_AMT.Read())
            {
                strVa = Convert.ToDouble(dReader_Today_AMT[0]);
            }

            dReader_Today_AMT.Close();
            SqlTransaction tran_Update_Trigger = conn.BeginTransaction();
            cmd.Transaction = tran_Update_Trigger;

            foreach (var item in dt)
            {
                try
                {
                    string sql = "update A_DAILY_AMT set dividend_last_amt = " + strVa + ",dividend_over_amt = " + strVa + "+" + Convert.ToDouble(item.DIVIDEND_AMT) + " where YYYY = '" + item.YYYY + "' AND MM = '" + item.MM + "' AND DD ='" + item.DD + "' ";

                    strVa = strVa + Convert.ToDouble(item.DIVIDEND_AMT);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    tran_Update_Trigger.Rollback();
                }
            }
            tran_Update_Trigger.Commit();
        }

        // 수정시 트리거 로직 (월자금 로직)
        protected void updt_trg_plan_value()
        {
            string amt_plan = "";
            DataSet1TableAdapters.USP_A_DAILY_AMT_PLAN_NEXTDAYTableAdapter adapter = new DataSet1TableAdapters.USP_A_DAILY_AMT_PLAN_NEXTDAYTableAdapter();
            DataSet1.USP_A_DAILY_AMT_PLAN_NEXTDAYDataTable dt = adapter.GetData(txt_yyyy.Text, txt_mm.Text, "01");
            // USP_A_DAILY_AMT_NEXTDAY 프로시저를 사용하여 해당 수정일 이후의 모든 데이터를 끌고온다 (전월, 금월, 차월 금액)

            amt_plan = " SELECT TERMEND_PAY FROM A_DAILY_AMT_PLAN WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "'AND DD = '01' ";
            // 수정된 차월데이터를 끌고와서 다음일 전월 데이터로 사용

            cmd.Connection = conn;
            cmd.CommandText = amt_plan;
            SqlDataReader dReader_amt_plan = cmd.ExecuteReader();

              double strVa = 0;

              if (dReader_amt_plan.Read())
            {
                strVa = Convert.ToDouble(dReader_amt_plan[0]);
            }

            dReader_amt_plan.Close();

            SqlTransaction tran_Update_Trigger = conn.BeginTransaction();
            cmd.Transaction = tran_Update_Trigger;

            foreach (var item in dt)
            {
                try
                {
                    string sql = "update A_DAILY_AMT_PLAN set BASIS_PAY = " + strVa + ", TERMEND_PAY = " + strVa + "+" + Convert.ToDouble(item.SPOTMONTH_PAY) + " where YYYY = '" + item.YYYY + "' AND MM = '" + item.MM + "' AND DD ='" + item.DD + "' ";

                    strVa = strVa + Convert.ToDouble(item.SPOTMONTH_PAY);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    tran_Update_Trigger.Rollback();
                }
            }
            tran_Update_Trigger.Commit();

        }

        // 수정시 트리거 로직 (월자금 실적로직)
        protected void updt_trg_plan_value_RSLT()
        {
            string amt_plan_rslt = "";
            DataSet1TableAdapters.USP_A_DAILY_AMT_PLAN_RSLT_NEXTDAYTableAdapter adapter = new DataSet1TableAdapters.USP_A_DAILY_AMT_PLAN_RSLT_NEXTDAYTableAdapter();
            DataSet1.USP_A_DAILY_AMT_PLAN_RSLT_NEXTDAYDataTable dt = adapter.GetData(txt_yyyy.Text, txt_mm.Text, "01");
            // USP_A_DAILY_AMT_NEXTDAY 프로시저를 사용하여 해당 수정일 이후의 모든 데이터를 끌고온다 (전월, 금월, 차월 금액)

            amt_plan_rslt = " SELECT TERMEND_PAY_RSLT FROM A_DAILY_AMT_PLAN WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "'AND DD = '01' ";
            // 수정된 차월데이터를 끌고와서 다음일 전월 데이터로 사용

            cmd.Connection = conn;
            cmd.CommandText = amt_plan_rslt;
            SqlDataReader dReader_amt_plan_rslt = cmd.ExecuteReader();

            double strVa = 0;

            if (dReader_amt_plan_rslt.Read())
            {
                strVa = Convert.ToDouble(dReader_amt_plan_rslt[0]);
            }

            dReader_amt_plan_rslt.Close();

            SqlTransaction tran_Update_Trigger = conn.BeginTransaction();
            cmd.Transaction = tran_Update_Trigger;

            foreach (var item in dt)
            {
                try
                {
                    string sql = "update A_DAILY_AMT_PLAN set BASIS_PAY_RSLT = " + strVa + ", TERMEND_PAY_RSLT = " + strVa + "+" + Convert.ToDouble(item.SPOTMONTH_PAY_RSLT) + " where YYYY = '" + item.YYYY + "' AND MM = '" + item.MM + "' AND DD ='" + item.DD + "' ";

                    strVa = strVa + Convert.ToDouble(item.SPOTMONTH_PAY_RSLT);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    tran_Update_Trigger.Rollback();
                }
            }
            tran_Update_Trigger.Commit();

        }

    }
}
