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
using ERPAppAddition.ERPAddition.AM.AM_A9001;

namespace ERPAppAddition.ERPAddition.AM.AM_A9001
{
    public partial class AM_A9001 : System.Web.UI.Page
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes_amc"].ConnectionString);
        SqlCommand cmd = new SqlCommand();
        string id = "";
        int select = 0;

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

        protected void calculation()
        {
            string LastMon = string.Empty;
            Double LastValue = 0;


            LastMon = "EXEC USP_A_DAILY_AMT_LASTMON '" + txt_yyyy.Text + "','" + txt_mm.Text + "', '" + txt_dd.Text + "'  ";

            cmd.Connection = conn;
            cmd.CommandText = LastMon;
            SqlDataReader dReader_LastMon = cmd.ExecuteReader();




            if (dReader_LastMon.Read())
            {
                LastValue = Int64.Parse(dReader_LastMon.GetValue(0).ToString());
            }
            
            else
            {
                LastValue = 0;
            }


            tb_amt7.Text = (Convert.ToDouble(tb_amt1.Text) + Convert.ToDouble(tb_amt2.Text) + Convert.ToDouble(tb_amt3.Text) + Convert.ToDouble(tb_amt4.Text)
                            + Convert.ToDouble(tb_amt5.Text) + Convert.ToDouble(tb_amt6.Text)).ToString(); // 8번 영업활동상의 자금수입 계산

            tb_amt13.Text = (Convert.ToDouble(tb_amt9.Text) + Convert.ToDouble(tb_amt10.Text) + Convert.ToDouble(tb_amt11.Text)
                             + Convert.ToDouble(tb_amt12.Text)).ToString(); // 14번 인건비의 지급 계산

            tb_amt19.Text = (Convert.ToDouble(tb_amt8.Text) + Convert.ToDouble(tb_amt13.Text) + Convert.ToDouble(tb_amt14.Text) + Convert.ToDouble(tb_amt15.Text)
                             + Convert.ToDouble(tb_amt16.Text) + Convert.ToDouble(tb_amt17.Text) + Convert.ToDouble(tb_amt18.Text)).ToString(); // 18번 영업상의 자금지출 계산

            tb_amt20.Text = (Convert.ToDouble(tb_amt7.Text) - Convert.ToDouble(tb_amt19.Text)).ToString(); // 20번 영업상의 순 자금유입 계산

            tb_amt25.Text = (Convert.ToDouble(tb_amt21.Text) + Convert.ToDouble(tb_amt22.Text) + Convert.ToDouble(tb_amt23.Text)
                             + Convert.ToDouble(tb_amt24.Text)).ToString(); // 25번 투자활동의 Cash in 계산

            tb_amt29.Text = (Convert.ToDouble(tb_amt26.Text) + Convert.ToDouble(tb_amt27.Text) + Convert.ToDouble(tb_amt28.Text)).ToString(); // 28번 고정자산의 취득 계산

            tb_amt33.Text = (Convert.ToDouble(tb_amt29.Text) + Convert.ToDouble(tb_amt30.Text) + Convert.ToDouble(tb_amt31.Text)
                             + Convert.ToDouble(tb_amt32.Text)).ToString(); // 33번 투자활동의 Cash out 계산

            tb_amt34.Text = (Convert.ToDouble(tb_amt25.Text) - Convert.ToDouble(tb_amt33.Text)).ToString(); // 34번 투자활동의 자금흐름 계산

            tb_amt38.Text = (Convert.ToDouble(tb_amt35.Text) + Convert.ToDouble(tb_amt36.Text) + Convert.ToDouble(tb_amt37.Text)).ToString(); // 38번 재무활동의 Cash in 계산

            tb_amt42.Text = (Convert.ToDouble(tb_amt39.Text) + Convert.ToDouble(tb_amt40.Text) + Convert.ToDouble(tb_amt41.Text)).ToString(); // 42번 재무활동의 Cash Out 계산

            tb_amt43.Text = (Convert.ToDouble(tb_amt38.Text) - Convert.ToDouble(tb_amt42.Text)).ToString(); // 43번 재무활동의 현금 흐름 계산

            tb_amt44.Text = (LastValue).ToString();

            tb_amt45.Text = (Convert.ToDouble(tb_amt20.Text) + Convert.ToDouble(tb_amt34.Text) + Convert.ToDouble(tb_amt43.Text)).ToString();

            tb_amt46.Text = (Convert.ToDouble(tb_amt44.Text) + Convert.ToDouble(tb_amt45.Text)).ToString();

            dReader_LastMon.Close();

        } // 계산로직


        protected void Update_Trigger() // data update 시 이후금액에 계산 및 반영되는 함수
        {
            conn.Open();     
            string Today_AMT = "";
            DataSet2TableAdapters.USP_A_DAILY_AMT_NEXTDAYTableAdapter adapter = new DataSet2TableAdapters.USP_A_DAILY_AMT_NEXTDAYTableAdapter();
            DataSet2.USP_A_DAILY_AMT_NEXTDAYDataTable dt = adapter.GetData(txt_yyyy.Text, txt_mm.Text, txt_dd.Text);
            // USP_A_DAILY_AMT_NEXTDAY 프로시저를 사용하여 해당 수정일 이후의 모든 데이터를 끌고온다 (전월, 금월, 차월 금액)

            Today_AMT = "SELECT DIVIDEND_OVER_AMT FROM A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "'AND DD = '" + txt_dd.Text + "' ";
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



        protected void btn_select_Click(object sender, EventArgs e) // 조회버튼 클릭
        {
            string select_queryStr;
            string[] temp;

            conn.Open();

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

            select_queryStr = "SELECT * FROM A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' and DD = '" + txt_dd.Text + "'   ";

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
            }
            
            
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

            btn_save.Enabled = false;

            conn.Close();
            dReader_select.Close();
            conn.Dispose();
        }


        protected void btn_save_Click(object sender, EventArgs e) // 저장버튼 클릭
        {

            select = 1; // 함수 식별자 (1 : 저장 / 2 : 업데이트)

            if (txt_yyyy.Text == "" ||
                txt_mm.Text == "" ||
                txt_dd.Text == "" ||
                tb_amt1.Text == "" ||
                tb_amt2.Text == "" ||
                tb_amt3.Text == "" ||
                tb_amt4.Text == "" ||
                tb_amt5.Text == "" ||
                tb_amt6.Text == "" ||
                //tb_amt7.Text == "" ||
                tb_amt8.Text == "" ||
                tb_amt9.Text == "" ||
                tb_amt10.Text == "" ||
                tb_amt11.Text == "" ||
                tb_amt12.Text == "" ||
                //tb_amt13.Text == "" ||
                tb_amt14.Text == "" ||
                tb_amt15.Text == "" ||
                tb_amt16.Text == "" ||
                tb_amt17.Text == "" ||
                //tb_amt18.Text == "" ||
                //tb_amt19.Text == "" ||
                //tb_amt20.Text == "" ||
                tb_amt21.Text == "" ||
                tb_amt22.Text == "" ||
                tb_amt23.Text == "" ||
                tb_amt24.Text == "" ||
                //tb_amt25.Text == "" ||
                tb_amt26.Text == "" ||
                tb_amt27.Text == "" ||
                tb_amt28.Text == "" ||
                //tb_amt29.Text == "" ||
                tb_amt30.Text == "" ||
                tb_amt31.Text == "" ||
                tb_amt32.Text == "" ||
                //tb_amt33.Text == "" ||
                //tb_amt34.Text == "" ||
                tb_amt35.Text == "" ||
                tb_amt36.Text == "" ||
                tb_amt37.Text == "" ||
                //tb_amt38.Text == "" ||
                tb_amt39.Text == "" ||
                tb_amt40.Text == "" ||
                tb_amt41.Text == "")
            //tb_amt42.Text == "" ||
            //tb_amt43.Text == "" ||
            //tb_amt44.Text == "" ||
            //tb_amt45.Text == "" ||
            //tb_amt46.Text == "")
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력되지 않은 항목이 있습니다.')", true);
            }
               
            else
            {
                
                
                conn.Open();
                cmd.Connection = conn;
                SqlTransaction tran_insert = conn.BeginTransaction();
                cmd.Transaction = tran_insert;

                string select_data = "";
                select_data = "SELECT * FROM A_DAILY_AMT WHERE YYYY+MM+DD >= '" + txt_yyyy.Text + "" + txt_mm.Text + "" + txt_dd.Text + "'";
                    
                cmd.CommandText = select_data;
                SqlDataReader dt_select = cmd.ExecuteReader();

                if (dt_select.Read())
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('저장을 시도한 날짜보다 이후 데이터가 존재합니다.\\n시스템관리자에게 문의바랍니다.')", true);
                    dt_select.Close(); // 데이터가 있는지 확인을 위한 리더를 닫는 문
                    return;
                }

                dt_select.Close(); // 데이터가 있는지 확인을 위한 리더를 닫는 문

                calculation(); // 계산함수 호출

                try
                {
                    string insert_queryStr;

                    insert_queryStr =
                    @"INSERT INTO [nepes_amc].[dbo].[A_DAILY_AMT]
                        ([YYYY]
                        ,[MM]
                        ,[DD]
                        ,[LENDER_COLLET]
                        ,[SURTAX_REFUND]
                        ,[TARIFF_REFUND]
                        ,[LEASE_INCOME]
                        ,[IMPORT_INCOME]
                        ,[ETC_1]
                        ,[ETC_1_BUSINESS_INCOME]
                        ,[M_MATERIAL]
                        ,[M_PAY]
                        ,[M_RETIRE]
                        ,[M_FOUNTAIN]
                        ,[M_WELRARE]
                        ,[PAYROLL_COSTS_EXPENSE]
                        ,[EXPENSE]
                        ,[LEASE_EXPENSE]
                        ,[SURTAX_PAYMENT]
                        ,[INSEREST]
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
                    insert_queryStr += "'',";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate(),";
                    insert_queryStr += "'" + id + "',";
                    insert_queryStr += "getdate())";

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
            }
            conn.Close();
        }



        //protected void btn_delete_Click(object sender, EventArgs e) // 삭제버튼 클릭
        //{
        //    conn.Open();

        //    string delete_queryStr;
        //    SqlTransaction tran_delete = conn.BeginTransaction();
        //    cmd.Connection = conn;
        //    cmd.Transaction = tran_delete;

        //    int intReturnRow_DELETE;

        //    try
        //    {
        //        delete_queryStr = "DELETE A_DAILY_AMT WHERE YYYY = '" + txt_yyyy.Text + "' AND MM = '" + txt_mm.Text + "' AND DD = '" + txt_dd.Text + "'";
        //        cmd.CommandText = delete_queryStr;
        //        intReturnRow_DELETE = cmd.ExecuteNonQuery();


        //        if (intReturnRow_DELETE == 0)
        //        {
        //            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('삭제할 데이터가 없습니다.')", true);
        //        }

        //        else
        //        {
        //            tran_delete.Commit();
        //            ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('삭제되었습니다.')", true);
        //        }
        //    }
        //    catch
        //    {
        //        tran_delete.Rollback();
        //        ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('애러가 발생되었습니다. 다시 시도하세요.')", true);
        //    }

        //    conn.Close();
        //} 
        // 삭제할 경우 이력관리가 되지 않기 때문에 수정버튼으로 대체 (05/11일 수정)

        protected void btn_reselect_Click(object sender, EventArgs e) //수정버튼 클릭
        {

            if (txt_yyyy.Text == "" ||
               txt_mm.Text == "" ||
               txt_dd.Text == "" ||
               tb_amt1.Text == "" ||
               tb_amt2.Text == "" ||
               tb_amt3.Text == "" ||
               tb_amt4.Text == "" ||
               tb_amt5.Text == "" ||
               tb_amt6.Text == "" ||
                //tb_amt7.Text == "" ||
               tb_amt8.Text == "" ||
               tb_amt9.Text == "" ||
               tb_amt10.Text == "" ||
               tb_amt11.Text == "" ||
               tb_amt12.Text == "" ||
                //tb_amt13.Text == "" ||
               tb_amt14.Text == "" ||
               tb_amt15.Text == "" ||
               tb_amt16.Text == "" ||
               tb_amt17.Text == "" ||
                //tb_amt18.Text == "" ||
                //tb_amt19.Text == "" ||
                //tb_amt20.Text == "" ||
               tb_amt21.Text == "" ||
               tb_amt22.Text == "" ||
               tb_amt23.Text == "" ||
               tb_amt24.Text == "" ||
                //tb_amt25.Text == "" ||
               tb_amt26.Text == "" ||
               tb_amt27.Text == "" ||
               tb_amt28.Text == "" ||
                //tb_amt29.Text == "" ||
               tb_amt30.Text == "" ||
               tb_amt31.Text == "" ||
               tb_amt32.Text == "" ||
                //tb_amt33.Text == "" ||
                //tb_amt34.Text == "" ||
               tb_amt35.Text == "" ||
               tb_amt36.Text == "" ||
               tb_amt37.Text == "" ||
                //tb_amt38.Text == "" ||
               tb_amt39.Text == "" ||
               tb_amt40.Text == "" ||
               tb_amt41.Text == "")
            //tb_amt42.Text == "" ||
            //tb_amt43.Text == "" ||
            //tb_amt44.Text == "" ||
            //tb_amt45.Text == "" ||
            //tb_amt46.Text == "")
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", "alert('입력되지 않은 항목이 있습니다.')", true);
            }

            else
            {

                conn.Open();

                string update_queryStr;
                SqlTransaction tran_update = conn.BeginTransaction();
                cmd.Connection = conn;
                cmd.Transaction = tran_update;

                int intReturnRow_UPDATE;

                select = 2; // 함수 식별자 (1 : 저장 / 2 : 업데이트)
                calculation(); // 계산함수 호출

                try
                {
                    update_queryStr = "UPDATE [nepes_amc].[dbo].[A_DAILY_AMT] SET ";

                    update_queryStr += "[LENDER_COLLET] = '" + Convert.ToDouble(tb_amt1.Text) + "',";
                    update_queryStr += "[SURTAX_REFUND] = '" + Convert.ToDouble(tb_amt2.Text) + "',";
                    update_queryStr += "[TARIFF_REFUND] = '" + Convert.ToDouble(tb_amt3.Text) + "',";
                    update_queryStr += "[LEASE_INCOME] = '" + Convert.ToDouble(tb_amt4.Text) + "',";
                    update_queryStr += "[IMPORT_INCOME] = '" + Convert.ToDouble(tb_amt5.Text) + "',";
                    update_queryStr += "[ETC_1] = '" + Convert.ToDouble(tb_amt6.Text) + "',";
                    update_queryStr += "[ETC_1_BUSINESS_INCOME] = '" + Convert.ToDouble(tb_amt7.Text) + "',";
                    update_queryStr += "[M_MATERIAL] = '" + Convert.ToDouble(tb_amt8.Text) + "',";
                    update_queryStr += "[M_PAY] = '" + Convert.ToDouble(tb_amt9.Text) + "',";
                    update_queryStr += "[M_RETIRE] = '" + Convert.ToDouble(tb_amt10.Text) + "',";
                    update_queryStr += "[M_FOUNTAIN] = '" + Convert.ToDouble(tb_amt11.Text) + "',";
                    update_queryStr += "[M_WELRARE] = '" + Convert.ToDouble(tb_amt12.Text) + "',";
                    update_queryStr += "[PAYROLL_COSTS_EXPENSE] = '" + Convert.ToDouble(tb_amt13.Text) + "',";
                    update_queryStr += "[EXPENSE] = '" + Convert.ToDouble(tb_amt14.Text) + "',";
                    update_queryStr += "[LEASE_EXPENSE] = '" + Convert.ToDouble(tb_amt15.Text) + "',";
                    update_queryStr += "[SURTAX_PAYMENT] = '" + Convert.ToDouble(tb_amt16.Text) + "',";
                    update_queryStr += "[INSEREST] = '" + Convert.ToDouble(tb_amt17.Text) + "',";
                    update_queryStr += "[ETC_2] = '" + Convert.ToDouble(tb_amt18.Text) + "',";
                    update_queryStr += "[ETC_2_BUSINESS_EXPENSE] = '" + Convert.ToDouble(tb_amt19.Text) + "',";
                    update_queryStr += "[ETC_2_BUSINESS_INFLOW] = '" + Convert.ToDouble(tb_amt20.Text) + "',";
                    update_queryStr += "[FIXED_ASSET_OUT] = '" + Convert.ToDouble(tb_amt21.Text) + "',";
                    update_queryStr += "[LOAN_INCOME] = '" + Convert.ToDouble(tb_amt22.Text) + "',";
                    update_queryStr += "[FINANCIAL_IN] = '" + Convert.ToDouble(tb_amt23.Text) + "',";
                    update_queryStr += "[VALUABLE_PAPAERS_INCOME] = '" + Convert.ToDouble(tb_amt24.Text) + "',";
                    update_queryStr += "[INVERST] = '" + Convert.ToDouble(tb_amt25.Text) + "',";
                    update_queryStr += "[INVEST_LAND] = '" + Convert.ToDouble(tb_amt26.Text) + "',";
                    update_queryStr += "[INVEST_MACHINE] = '" + Convert.ToDouble(tb_amt27.Text) + "',";
                    update_queryStr += "[INVEST_CAR] = '" + Convert.ToDouble(tb_amt28.Text) + "',";
                    update_queryStr += "[FIXED_ASSET_IN] = '" + Convert.ToDouble(tb_amt29.Text) + "',";
                    update_queryStr += "[LOAN_EXPENSE] = '" + Convert.ToDouble(tb_amt30.Text) + "',";
                    update_queryStr += "[FINANCIAL_OUT] = '" + Convert.ToDouble(tb_amt31.Text) + "',";
                    update_queryStr += "[INVERST_OUT] = '" + Convert.ToDouble(tb_amt32.Text) + "',";
                    update_queryStr += "[INVEST_OUT_CASHOUT] = '" + Convert.ToDouble(tb_amt33.Text) + "',";
                    update_queryStr += "[INVEST_OUT_CASH] = '" + Convert.ToDouble(tb_amt34.Text) + "',";
                    update_queryStr += "[BEWBORROW] = '" + Convert.ToDouble(tb_amt35.Text) + "',";
                    update_queryStr += "[USANCE_CD1] = '" + Convert.ToDouble(tb_amt36.Text) + "',";
                    update_queryStr += "[INCRES_CAPITAL] = '" + Convert.ToDouble(tb_amt37.Text) + "',";
                    update_queryStr += "[INCRES_CAPITAL_CASH] = '" + Convert.ToDouble(tb_amt38.Text) + "',";
                    update_queryStr += "[LOAN_CD1] = '" + Convert.ToDouble(tb_amt39.Text) + "',";
                    update_queryStr += "[USANCE_CD2] = '" + Convert.ToDouble(tb_amt40.Text) + "',";
                    update_queryStr += "[DIVIDEND] = '" + Convert.ToDouble(tb_amt41.Text) + "',";
                    update_queryStr += "[DIVIDEND_CASHOUT] = '" + Convert.ToDouble(tb_amt42.Text) + "',";
                    update_queryStr += "[DIVIDEND_CASH] = '" + Convert.ToDouble(tb_amt43.Text) + "',";
                    update_queryStr += "[DIVIDEND_LAST_AMT] = '" + Convert.ToDouble(tb_amt44.Text) + "',";
                    update_queryStr += "[DIVIDEND_AMT] = '" + Convert.ToDouble(tb_amt45.Text) + "',";
                    update_queryStr += "[DIVIDEND_OVER_AMT] = '" + Convert.ToDouble(tb_amt46.Text) + "',";
                    //update_queryStr += "[INSRT_USER_ID] = '',";
                    update_queryStr += "[INSRT_DT] = getdate(),";
                    update_queryStr += "[UPDT_USER_ID] = '" + id + "',";
                    update_queryStr += "[UPDT_DT] = getdate()";
                    update_queryStr += "WHERE YYYY = '" + txt_yyyy.Text + "' and  MM = '" + txt_mm.Text + "' and DD = '" + txt_dd.Text + "'   ";

                    cmd.CommandText = update_queryStr;
                    intReturnRow_UPDATE = cmd.ExecuteNonQuery();

                    if (intReturnRow_UPDATE == 0)
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
                conn.Close();
            }
            Update_Trigger();
        }
    }
     
}




    









