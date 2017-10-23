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
using System.Collections.Generic;

namespace ERPAppAddition.ERPAddition.SM.sm_s9022
{
    public partial class sm_s9022c : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        //string sql_cust_cd;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand();
        DataSet ds = new DataSet();

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleDataReader dr;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                /*달력셋*/
                setMonth();
                WebSiteCount();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = sql_conn.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void setMonth()
        {
            conn.Open();
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("select PLAN_YEAR || lpad(PLAN_MONTH, 2, 0) AS M_MONTH   \n");
                sbSQL.Append("  from CALENDAR                                         \n");
                sbSQL.Append(" where PLANT = 'CCUBEDIGITAL'                           \n");
                sbSQL.Append("   and PLAN_YEAR >= '2015'                              \n");
                sbSQL.Append("   and PLAN_YEAR <= to_char(sysdate, 'yyyy')            \n");
                sbSQL.Append("   and PLAN_YEAR || LPAD(PLAN_MONTH, 2, 0) <= to_char(sysdate-1, 'YYYYMM') \n");
                sbSQL.Append(" group by PLAN_YEAR, PLAN_MONTH                         \n");
                sbSQL.Append(" order by 1 desc                                             \n");
                OracleCommand cmd2 = new OracleCommand(sbSQL.ToString(), conn);

                dr = cmd2.ExecuteReader();

                if (dr.RowSize > 0)
                {
                    tb_yyyymm.DataSource = dr;
                    tb_yyyymm.DataValueField = "M_MONTH";
                    tb_yyyymm.DataTextField = "M_MONTH";
                    tb_yyyymm.DataBind();
                    tb_yyyymm.SelectedValue = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00");
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            //return dr;
        }



        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            if (tb_yyyymm.Text == "" || tb_yyyymm.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"조회년을 선택해주세요.\");";
                ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
            }
            else
            {
                setGrid();
            }
        }

        private void setGrid()
        {
            try
            {
                // 프로시져 실행: 기본데이타 생성
                sql_conn.Open();
                sql_cmd = sql_conn.CreateCommand();
                sql_cmd.CommandType = CommandType.Text;
                sql_cmd.CommandText = getSQL();
                sql_cmd.CommandTimeout = 0;

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd);

                    da.Fill(ds, "DataSet1");
                }
                catch (Exception ex)
                {
                    if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                    //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
                }
                sql_conn.Close();

                /*seq 가 a는 필수 항목 없는경우 조회 불가*/
                if (ds.Tables["DataSet1"].Rows.Count <= 0)
                {
                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s9022c.rdlc");
                ReportViewer1.LocalReport.DisplayName = "재고출고관리 조회" + DateTime.Now.ToShortDateString();

                string[] itemCd = { "1B30-R0000", "1B30-R0002", "1B30-R0003", "1B30-R0004", "1B30-R0005", "1B30-R0100", "1B30-R0101", "1B30-R0102", "1B30-R0103", "1B35-R0000","1F35-R0000"};


                ReportDataSource rdsRow1_1 = new ReportDataSource(); ReportDataSource rdsRow2_1 = new ReportDataSource(); ReportDataSource rdsRow3_1 = new ReportDataSource();
                ReportDataSource rdsRow4_1 = new ReportDataSource(); ReportDataSource rdsRow5_1 = new ReportDataSource(); ReportDataSource rdsRow6_1 = new ReportDataSource(); ReportDataSource rdsRow7_1 = new ReportDataSource();

                ReportDataSource rdsRow1_2 = new ReportDataSource(); ReportDataSource rdsRow2_2 = new ReportDataSource(); ReportDataSource rdsRow3_2 = new ReportDataSource();
                ReportDataSource rdsRow4_2 = new ReportDataSource(); ReportDataSource rdsRow5_2 = new ReportDataSource(); ReportDataSource rdsRow6_2 = new ReportDataSource(); ReportDataSource rdsRow7_2 = new ReportDataSource();

                ReportDataSource rdsRow1_3 = new ReportDataSource(); ReportDataSource rdsRow2_3 = new ReportDataSource(); ReportDataSource rdsRow3_3 = new ReportDataSource();
                ReportDataSource rdsRow4_3 = new ReportDataSource(); ReportDataSource rdsRow5_3 = new ReportDataSource(); ReportDataSource rdsRow6_3 = new ReportDataSource(); ReportDataSource rdsRow7_3 = new ReportDataSource();

                ReportDataSource rdsRow1_4 = new ReportDataSource(); ReportDataSource rdsRow2_4 = new ReportDataSource(); ReportDataSource rdsRow3_4 = new ReportDataSource();
                ReportDataSource rdsRow4_4 = new ReportDataSource(); ReportDataSource rdsRow5_4 = new ReportDataSource(); ReportDataSource rdsRow6_4 = new ReportDataSource(); ReportDataSource rdsRow7_4 = new ReportDataSource();

                ReportDataSource rdsRow1_5 = new ReportDataSource(); ReportDataSource rdsRow2_5 = new ReportDataSource(); ReportDataSource rdsRow3_5 = new ReportDataSource();
                ReportDataSource rdsRow4_5 = new ReportDataSource(); ReportDataSource rdsRow5_5 = new ReportDataSource(); ReportDataSource rdsRow6_5 = new ReportDataSource(); ReportDataSource rdsRow7_5 = new ReportDataSource();

                ReportDataSource rdsRow1_6 = new ReportDataSource(); ReportDataSource rdsRow2_6 = new ReportDataSource(); ReportDataSource rdsRow3_6 = new ReportDataSource();
                ReportDataSource rdsRow4_6 = new ReportDataSource(); ReportDataSource rdsRow5_6 = new ReportDataSource(); ReportDataSource rdsRow6_6 = new ReportDataSource(); ReportDataSource rdsRow7_6 = new ReportDataSource();

                ReportDataSource rdsRow1_7 = new ReportDataSource(); ReportDataSource rdsRow2_7 = new ReportDataSource(); ReportDataSource rdsRow3_7 = new ReportDataSource();
                ReportDataSource rdsRow4_7 = new ReportDataSource(); ReportDataSource rdsRow5_7 = new ReportDataSource(); ReportDataSource rdsRow6_7 = new ReportDataSource(); ReportDataSource rdsRow7_7 = new ReportDataSource();

                ReportDataSource rdsRow1_8 = new ReportDataSource(); ReportDataSource rdsRow2_8 = new ReportDataSource(); ReportDataSource rdsRow3_8 = new ReportDataSource();
                ReportDataSource rdsRow4_8 = new ReportDataSource(); ReportDataSource rdsRow5_8 = new ReportDataSource(); ReportDataSource rdsRow6_8 = new ReportDataSource(); ReportDataSource rdsRow7_8 = new ReportDataSource();

                ReportDataSource rdsRow1_9 = new ReportDataSource(); ReportDataSource rdsRow2_9 = new ReportDataSource(); ReportDataSource rdsRow3_9 = new ReportDataSource();
                ReportDataSource rdsRow4_9 = new ReportDataSource(); ReportDataSource rdsRow5_9 = new ReportDataSource(); ReportDataSource rdsRow6_9 = new ReportDataSource(); ReportDataSource rdsRow7_9 = new ReportDataSource();

                ReportDataSource rdsRow1_10 = new ReportDataSource(); ReportDataSource rdsRow2_10 = new ReportDataSource(); ReportDataSource rdsRow3_10 = new ReportDataSource();
                ReportDataSource rdsRow4_10 = new ReportDataSource(); ReportDataSource rdsRow5_10 = new ReportDataSource(); ReportDataSource rdsRow6_10 = new ReportDataSource(); ReportDataSource rdsRow7_10 = new ReportDataSource();

                ReportDataSource rdsRow1_11 = new ReportDataSource(); ReportDataSource rdsRow2_11 = new ReportDataSource(); ReportDataSource rdsRow3_11 = new ReportDataSource();
                ReportDataSource rdsRow4_11 = new ReportDataSource(); ReportDataSource rdsRow5_11 = new ReportDataSource(); ReportDataSource rdsRow6_11 = new ReportDataSource(); ReportDataSource rdsRow7_11 = new ReportDataSource();
                
                Dictionary<string, ReportDataSource> dic = new Dictionary<string, ReportDataSource>();
                dic.Add("rdsRow1_1", rdsRow1_1); dic.Add("rdsRow2_1", rdsRow2_1); dic.Add("rdsRow3_1", rdsRow3_1); dic.Add("rdsRow4_1", rdsRow4_1);
                dic.Add("rdsRow5_1", rdsRow5_1); dic.Add("rdsRow6_1", rdsRow6_1); dic.Add("rdsRow7_1", rdsRow7_1);
                
                dic.Add("rdsRow1_2", rdsRow1_2); dic.Add("rdsRow2_2", rdsRow2_2); dic.Add("rdsRow3_2", rdsRow3_2); dic.Add("rdsRow4_2", rdsRow4_2);
                dic.Add("rdsRow5_2", rdsRow5_2); dic.Add("rdsRow6_2", rdsRow6_2); dic.Add("rdsRow7_2", rdsRow7_2);

                dic.Add("rdsRow1_3", rdsRow1_3); dic.Add("rdsRow2_3", rdsRow2_3); dic.Add("rdsRow3_3", rdsRow3_3); dic.Add("rdsRow4_3", rdsRow4_3);
                dic.Add("rdsRow5_3", rdsRow5_3); dic.Add("rdsRow6_3", rdsRow6_3); dic.Add("rdsRow7_3", rdsRow7_3);

                dic.Add("rdsRow1_4", rdsRow1_4); dic.Add("rdsRow2_4", rdsRow2_4); dic.Add("rdsRow3_4", rdsRow3_4); dic.Add("rdsRow4_4", rdsRow4_4);
                dic.Add("rdsRow5_4", rdsRow5_4); dic.Add("rdsRow6_4", rdsRow6_4); dic.Add("rdsRow7_4", rdsRow7_4);

                dic.Add("rdsRow1_5", rdsRow1_5); dic.Add("rdsRow2_5", rdsRow2_5); dic.Add("rdsRow3_5", rdsRow3_5); dic.Add("rdsRow4_5", rdsRow4_5);
                dic.Add("rdsRow5_5", rdsRow5_5); dic.Add("rdsRow6_5", rdsRow6_5); dic.Add("rdsRow7_5", rdsRow7_5);

                dic.Add("rdsRow1_6", rdsRow1_6); dic.Add("rdsRow2_6", rdsRow2_6); dic.Add("rdsRow3_6", rdsRow3_6); dic.Add("rdsRow4_6", rdsRow4_6);
                dic.Add("rdsRow5_6", rdsRow5_6); dic.Add("rdsRow6_6", rdsRow6_6); dic.Add("rdsRow7_6", rdsRow7_6);

                dic.Add("rdsRow1_7", rdsRow1_7); dic.Add("rdsRow2_7", rdsRow2_7); dic.Add("rdsRow3_7", rdsRow3_7); dic.Add("rdsRow4_7", rdsRow4_7);
                dic.Add("rdsRow5_7", rdsRow5_7); dic.Add("rdsRow6_7", rdsRow6_7); dic.Add("rdsRow7_7", rdsRow7_7);

                dic.Add("rdsRow1_8", rdsRow1_8); dic.Add("rdsRow2_8", rdsRow2_8); dic.Add("rdsRow3_8", rdsRow3_8); dic.Add("rdsRow4_8", rdsRow4_8);
                dic.Add("rdsRow5_8", rdsRow5_8); dic.Add("rdsRow6_8", rdsRow6_8); dic.Add("rdsRow7_8", rdsRow7_8);

                dic.Add("rdsRow1_9", rdsRow1_9); dic.Add("rdsRow2_9", rdsRow2_9); dic.Add("rdsRow3_9", rdsRow3_9); dic.Add("rdsRow4_9", rdsRow4_9);
                dic.Add("rdsRow5_9", rdsRow5_9); dic.Add("rdsRow6_9", rdsRow6_9); dic.Add("rdsRow7_9", rdsRow7_9);

                dic.Add("rdsRow1_10", rdsRow1_10); dic.Add("rdsRow2_10", rdsRow2_10); dic.Add("rdsRow3_10", rdsRow3_10); dic.Add("rdsRow4_10", rdsRow4_10);
                dic.Add("rdsRow5_10", rdsRow5_10); dic.Add("rdsRow6_10", rdsRow6_10); dic.Add("rdsRow7_10", rdsRow7_10);

                dic.Add("rdsRow1_11", rdsRow1_11); dic.Add("rdsRow2_11", rdsRow2_11); dic.Add("rdsRow3_11", rdsRow3_11); dic.Add("rdsRow4_11", rdsRow4_11);
                dic.Add("rdsRow5_11", rdsRow5_11); dic.Add("rdsRow6_11", rdsRow6_11); dic.Add("rdsRow7_11", rdsRow7_11);

                for (int i = 0; i < itemCd.Length; i++) //itemCd.Length
                {
                    DataTable dt1 = ds.Tables["DataSet1"].Copy();
                    dt1.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('공정실적','재공수량','환산수량')";
                    dic["rdsRow1_"+ (i+1).ToString()].Name = "DataSet1" + "_" + (i + 1).ToString();
                    dic["rdsRow1_" + (i + 1).ToString()].Value = dt1.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow1_" + (i + 1).ToString()]);
                    
                    DataTable dt2 = ds.Tables["DataSet1"].Copy();
                    dt2.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('Usage(BOM)')";
                    dic["rdsRow2_" + (i + 1).ToString()].Name = "DataSet2" + "_" + (i + 1).ToString(); ;
                    dic["rdsRow2_" + (i + 1).ToString()].Value = dt2.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow2_" + (i + 1).ToString()]);
                    
                    DataTable dt3 = ds.Tables["DataSet1"].Copy();
                    dt3.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('생산라인 투입수량')";
                    dic["rdsRow3_" + (i + 1).ToString()].Name = "DataSet3" + "_" + (i + 1).ToString(); ; ;
                    dic["rdsRow3_" + (i + 1).ToString()].Value = dt3.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow3_" + (i + 1).ToString()]);
                    
                    DataTable dt4 = ds.Tables["DataSet1"].Copy();
                    dt4.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('BOM대비 투입수량비율')";
                    dic["rdsRow4_" + (i + 1).ToString()].Name = "DataSet4" + "_" + (i + 1).ToString(); ; ;
                    dic["rdsRow4_" + (i + 1).ToString()].Value = dt4.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow4_" + (i + 1).ToString()]);
                    
                    DataTable dt5 = ds.Tables["DataSet1"].Copy();
                    dt5.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('이론재료비', '투입재료비')";
                    dic["rdsRow5_" + (i + 1).ToString()].Name = "DataSet5" + "_" + (i + 1).ToString(); ; ;
                    dic["rdsRow5_" + (i + 1).ToString()].Value = dt5.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow5_" + (i + 1).ToString()]);
                    
                    DataTable dt6 = ds.Tables["DataSet1"].Copy();
                    dt6.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('BOM 대비 투입금액비율')";
                    dic["rdsRow6_" + (i + 1).ToString()].Name = "DataSet6" + "_" + (i + 1).ToString(); ; ;
                    dic["rdsRow6_" + (i + 1).ToString()].Value = dt6.DefaultView;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow6_" + (i + 1).ToString()]);

                    /*그래프*/                    
                    DataTable dt7 = ds.Tables["DataSet1"].Copy();
                    dt7.DefaultView.RowFilter = "ITEM_CD = '" + itemCd[i].ToString() + "' AND TIT IN('환산수량', 'Usage(BOM)', '생산라인 투입수량')";
                    DataTable rowDt7 = dt7.DefaultView.ToTable();

                    DataTable dtCol7 = new DataTable();
                    dtCol7.Columns.Add("DT");
                    dtCol7.Columns.Add("CONV_UNIT");
                    dtCol7.Columns.Add("USAGE_BOM");
                    dtCol7.Columns.Add("INPUT_UNIT");

                    for (int j = 0; j < 31; j++)
                    {
                        DataRow dr = dtCol7.NewRow();
                        dr[0] = (j+1).ToString("00");
                        dr[1] = rowDt7.Rows[0]["T" + (j + 1).ToString("00")];
                        dr[2] = rowDt7.Rows[1]["T" + (j + 1).ToString("00")];
                        dr[3] = rowDt7.Rows[2]["T" + (j + 1).ToString("00")];

                        dtCol7.Rows.Add(dr);
                    }

                    dic["rdsRow7_" + (i + 1).ToString()].Name = "DataSet7" + "_" + (i + 1).ToString(); ; ;
                    dic["rdsRow7_" + (i + 1).ToString()].Value = dtCol7;
                    ReportViewer1.LocalReport.DataSources.Add(dic["rdsRow7_" + (i + 1).ToString()]);
                }


                ReportViewer1.ShowRefreshButton = false;  //새로고침 단추표시 x
            }
            catch (Exception ex)
            {
                if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
            }
        }

        private string getSQL()
        {
            string strDATE = tb_yyyymm.Text;

            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("select SEQ, ITEM_CD, CITEM_NM, CITEM_ITEM_UNIT as CHILD_ITEM_UNIT, TIT, TOTAL  \n");
            sbSQL.Append("      ,T01 ,T02 ,T03 ,T04 ,T05 ,T06 ,T07 ,T08 ,T09 ,T10  \n");
            sbSQL.Append("      ,T11 ,T12 ,T13 ,T14 ,T15 ,T16 ,T17 ,T18 ,T19 ,T20  \n");
            sbSQL.Append("      ,T01 ,T02 ,T03 ,T04 ,T05 ,T06 ,T07 ,T08 ,T09 ,T10  \n");
            sbSQL.Append("      ,T21 ,T22 ,T23 ,T24 ,T25 ,T26 ,T27 ,T28 ,T29 ,T30 ,T31  \n");
            sbSQL.Append(" FROM DBO.T_USP_S_BOM_CHK_SUM  \n");
            sbSQL.Append("WHERE NUM  = (SELECT MAX(NUM) FROM T_USP_S_BOM_CHK_SUM  \n");
            sbSQL.Append("                             WHERE DT = '" + strDATE + "'         \n");
            sbSQL.Append("              )                                         \n");
            sbSQL.Append(" AND DT = '" + strDATE + "'                                      \n");
            sbSQL.Append(" ORDER BY ITEM_CD, SEQ                                  \n");
            //sbSQL.Append("select 'aaa' as item_cd  \n");
            return sbSQL.ToString();
        }
    }
}