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

namespace ERPAppAddition.ERPAddition.SM.sm_s9011
{
    public partial class sm_s9011 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];
        //string sql_cust_cd;

        SqlConnection sql_conn = new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);
        //SqlConnection sql_conn1 = new SqlConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        SqlCommand sql_cmd = new SqlCommand(),  sql_cmd2 = new SqlCommand(), sql_cmd3 = new SqlCommand(),  sql_cmd5 = new SqlCommand();
        SqlCommand sql_cmd6 = new SqlCommand();
        SqlCommand sql_cmd7 = new SqlCommand();

        //SqlDataReader sql_dr;
        DataSet ds = new DataSet();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                tb_fr_yyyymmdd.Text = DateTime.Today.AddDays(-1).Year.ToString("0000") + DateTime.Today.AddDays(-1).Month.ToString("00") + "01";
                tb_to_yyyymmdd.Text = DateTime.Today.AddDays(-1).Year.ToString("0000") + DateTime.Today.AddDays(-1).Month.ToString("00") + DateTime.Today.AddDays(-1).Day.ToString("00");
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

        protected void Button1_Click(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
            if (tb_fr_yyyymmdd.Text == "" || tb_fr_yyyymmdd.Text == null || tb_to_yyyymmdd.Text == "" || tb_to_yyyymmdd.Text == null)
            {
                //MessageBox.ShowMessage("조회년도를 입력해주세요.", this.Page);
                string script = "alert(\"Date 를 선택해주세요.\");";
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

                sql_cmd2 = sql_conn.CreateCommand();
                sql_cmd2.CommandType = CommandType.Text;
                sql_cmd2.CommandText = getSQL2();

                sql_cmd3 = sql_conn.CreateCommand();
                sql_cmd3.CommandType = CommandType.Text;
                sql_cmd3.CommandText = getSQL3();

                sql_cmd5 = sql_conn.CreateCommand();
                sql_cmd5.CommandType = CommandType.Text;
                sql_cmd5.CommandText = getSQL5();

                sql_cmd6 = sql_conn.CreateCommand();
                sql_cmd6.CommandType = CommandType.Text;
                sql_cmd6.CommandText = getSQL6();

                sql_cmd7 = sql_conn.CreateCommand();
                sql_cmd7.CommandType = CommandType.Text;
                sql_cmd7.CommandText = getSQL7();

                DataTable dt = new DataTable();
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_cmd),  da2 = new SqlDataAdapter(sql_cmd2);
                    SqlDataAdapter da3 = new SqlDataAdapter(sql_cmd3),  da5 = new SqlDataAdapter(sql_cmd5);
                    SqlDataAdapter da6 = new SqlDataAdapter(sql_cmd6);
                    SqlDataAdapter da7 = new SqlDataAdapter(sql_cmd7);
                    
                    da.Fill(ds, "DataSet1");
                    da2.Fill(ds, "DataSet3");
                    da3.Fill(ds, "DataSet2");
                    da5.Fill(ds, "DataSet5");
                    da6.Fill(ds, "DataSet6");
                    da7.Fill(ds, "DataSet7");                    
                }
                catch (Exception ex)
                {                    
                    if (sql_conn.State == ConnectionState.Open) sql_conn.Close();
                    //if (sql_conn1.State == ConnectionState.Open) sql_conn1.Close();
                }
                sql_conn.Close();

                /*seq 가 a는 필수 항목 없는경우 조회 불가*/                
                if (ds.Tables["DataSet1"].Rows.Count <= 0 )
                {
                    string script = "alert(\"조회된 데이터가 없습니다..\");";
                    ScriptManager.RegisterStartupScript(this, GetType(), "ServerControlScript", script, true);
                    return;
                }

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("rd_sm_s9011.rdlc");
                ReportViewer1.LocalReport.DisplayName = "일별 손익레포트" + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSet1";
                DataTable dt1 = ds.Tables["DataSet1"].Copy();/*판관비를 본부비 로 수정*/
                dt1.DefaultView.RowFilter = "GUBN1 NOT IN('본부비','영업외손익','경상이익','전일 인당매출액','전일 인당영업이익','인당매출액','인당영업이익') ";
                rds.Value = dt1.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds);

                ReportDataSource rds10 = new ReportDataSource();
                rds10.Name = "DataSet10";
                DataTable dt10 = ds.Tables["DataSet1"].Copy();
                dt10.DefaultView.RowFilter = "GUBN1 IN('경상이익','영업외손익','본부비') ";
                rds10.Value = dt10.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds10);



                ReportDataSource rds2 = new ReportDataSource();
                rds2.Name = "DataSet2";
                rds2.Value = ds.Tables["DataSet2"];
                ReportViewer1.LocalReport.DataSources.Add(rds2);

                /*그래프*/
                ReportDataSource rds3 = new ReportDataSource();
                rds3.Name = "DataSet3";
                rds3.Value = ds.Tables["DataSet3"];
                ReportViewer1.LocalReport.DataSources.Add(rds3);

                /*추가된 인당 영업비용*/
                ReportDataSource rds4 = new ReportDataSource();
                rds4.Name = "DataSet4";
                DataTable dt4 = ds.Tables["DataSet1"].Copy();
                dt4.DefaultView.RowFilter = "GUBN1 IN('인당매출액','인당영업이익')";
                rds4.Value = dt4.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds4);

                /*박기수요청 20151119 전일 추정손익상세*/
                ReportDataSource rds5 = new ReportDataSource();
                rds5.Name = "DataSet5";
                DataTable dt5 = ds.Tables["DataSet5"].Copy();
                dt5.DefaultView.RowFilter = "GUBN1 NOT IN('본부비','영업외손익','경상이익','인당매출액','인당영업이익','전일 인당매출액','전일 인당영업이익') ";
                rds5.Value = dt5.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds5);

                ReportDataSource rds11 = new ReportDataSource();
                rds11.Name = "DataSet11";
                DataTable dt11 = ds.Tables["DataSet5"].Copy();
                dt11.DefaultView.RowFilter = "GUBN1 IN('경상이익','영업외손익','본부비') ";
                rds11.Value = dt11.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds11);


                ReportDataSource rds6 = new ReportDataSource();
                rds6.Name = "DataSet6";
                DataTable dt6 = ds.Tables["DataSet6"].Copy();
                rds6.Value = dt6.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds6);

                ReportDataSource rds7 = new ReportDataSource();
                rds7.Name = "DataSet7";
                DataTable dt7 = ds.Tables["DataSet7"].Copy();
                rds7.Value = dt7.DefaultView;
                ReportViewer1.LocalReport.DataSources.Add(rds7);

                
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
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_ACCOUNT  '" + strFrom + "', '" + strTo  + "' \n");
            return sbSQL.ToString();
        }

        private string getSQL2()
        {
            string strValue = tb_fr_yyyymmdd.Text;   
            /* 그래프 조회쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_ACCOUNT_MONTH  '" + strValue + "'\n");
            return sbSQL.ToString();
        }
        private string getSQL3()
        {
            string strFrom = tb_fr_yyyymmdd.Text;
            string strTo = tb_to_yyyymmdd.Text;
            /* 실적 조회 쿼리*/
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_ACCOUNT_DET  '" + strFrom + "', '" + strTo + "' \n");
            return sbSQL.ToString();
        }
        /*전일용으로 추가 요청 박기수20151119*/
        private string getSQL5()
        {
            string date = tb_to_yyyymmdd.Text.Substring(0, 4) + "-" + tb_to_yyyymmdd.Text.Substring(4, 2) + "-" + tb_to_yyyymmdd.Text.Substring(6, 2);
            /* 실적 조회 쿼리*/
            DateTime bfday = Convert.ToDateTime(date);
            string str = bfday.Year.ToString("0000") + bfday.Month.ToString("00") + bfday.Day.ToString("00");

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_S_ACCOUNT  '" + str + "', '" + str + "' \n");
            return sbSQL.ToString();
        }
        // 카테고리별 재공/재고
        private string getSQL6()
        {
            string strToDate = tb_to_yyyymmdd.Text, strSQL;

            strSQL = "SELECT * FROM OPENQUERY([CCUBE], 'WITH PWPT AS ( \n";
            strSQL = strSQL + "    SELECT CASE WHEN OPERATION IN (''5000'', ''8900'', ''9000'', ''FS40'', ''FS90'') THEN ''INV'' ELSE ''WIP'' END WIPINV, PART, QTY_UNIT_1, \n";
            strSQL = strSQL + "           OPERATION, PT.PROC_TYPE, PROD_TYPE, SUB_PLANT_1, SUB_PLANT_2, SUM(QTY_1) QTY, ''" + strToDate + "'' WORK_DATE \n";
            strSQL = strSQL + "      FROM POINTWIP@RPTMIT PW, PROC_TYPE_INFO PT \n";
            strSQL = strSQL + "     WHERE PW.PLANT = ''CCUBEDIGITAL'' AND PW.PLANT = PT.PLANT AND PW.CREATE_CODE = PT.PROC_TYPE AND PW.REWORK = ''N'' \n";
            strSQL = strSQL + "       AND PW.OPERATION BETWEEN PT.IN_OPER AND PT.OUT_OPER2 AND PW.STATUS <> 99 AND PW.PART_TYPE IN (''D'', ''P'') \n";
            strSQL = strSQL + "       AND PW.POINT_TIME = ''" + strToDate + "07'' \n";
            strSQL = strSQL + "     GROUP BY PART, OPERATION, PT.PROC_TYPE, PROD_TYPE, SUB_PLANT_1, SUB_PLANT_2, QTY_UNIT_1 \n";
            strSQL = strSQL + ") \n";
            strSQL = strSQL + "SELECT TO_DATE(''" + strToDate + "'') WORK_DATE, WIPINV, CATG1, CATG2, ROUND(SUM(CST)/1000000, 0) SCOST, SUM(QTY) SQTY \n";
            strSQL = strSQL + "  FROM ( \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''DDI'' CATG1, CASE WHEN PROD_TYPE = ''C'' THEN ''12'' ELSE ''8'' END || ''\"DDI_BUMP'' CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''ALL'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE, PARTSPEC PS  \n";
            strSQL = strSQL + "     WHERE PS.PLANT = ''CCUBEDIGITAL'' AND PP.PART = PS.PART_ID AND PP.SUB_PLANT_1 IN (''BUMP'', ''12BUMP'') AND SUB_PLANT_2 = ''DDI'' \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''DDI'' CATG1, ''DDI_TEST'' CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''ALL'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE \n";
            strSQL = strSQL + "     WHERE PP.SUB_PLANT_1 = ''P-TEST'' AND SUB_PLANT_2 = ''DDI'' \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''DDI'' CATG1, CASE WHEN PP.SUB_PLANT_1 = ''TAB'' THEN CASE WHEN OPERATION < ''7000'' THEN ''ASSY_COF'' ELSE ''FT_COF'' END ELSE ''COG'' END CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY  \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''DDI'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE \n";
            strSQL = strSQL + "     WHERE PP.SUB_PLANT_1 IN (''TAB'', ''COG'') \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''WLP'' CATG1, CASE PROD_TYPE WHEN ''C'' THEN ''12'' ELSE ''8'' END || ''\"WLP_BUMP'' CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''ALL'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE, PARTSPEC PS \n";
            strSQL = strSQL + "     WHERE PS.PLANT = ''CCUBEDIGITAL'' AND PP.PART = PS.PART_ID AND PP.SUB_PLANT_1 IN (''BUMP'', ''12BUMP'') AND SUB_PLANT_2 = ''WLP'' \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''WLP'' CATG1, ''WLP_TEST'' CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''ALL'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE \n";
            strSQL = strSQL + "     WHERE PP.SUB_PLANT_1 = ''P-TEST'' AND SUB_PLANT_2 = ''WLP'' \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''WLP'' CATG1, ''WLP_BE'' CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''WLP'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE \n";
            strSQL = strSQL + "     WHERE PP.SUB_PLANT_1 = ''WLCSP'' \n";
            strSQL = strSQL + "     UNION ALL \n";
            strSQL = strSQL + "    SELECT PP.WIPINV, ''FOWLP'' CATG1, CASE PP.SUB_PLANT_1 WHEN ''RCP-PANEL'' THEN ''FOWLP_PANEL'' ELSE ''FOWLP_PKG'' END CATG2, QTY * NVL(OC.UNIT_COST, 0) CST, QTY \n";
            strSQL = strSQL + "      FROM PWPT PP LEFT JOIN OPRUNTCST@RPTMIT OC ON PP.OPERATION = OC.OPER AND OC.CREATE_CODE = ''FOWLP'' \n";
            strSQL = strSQL + "       AND PP.WORK_DATE BETWEEN OC.YYYYMMDD AND OC.EXPIRY_DATE \n";
            strSQL = strSQL + "     WHERE PP.SUB_PLANT_1 IN (''RCP-PANEL'', ''RCP-FINAL'') \n";
            strSQL = strSQL + " ) \n";
            strSQL = strSQL + " GROUP BY WIPINV, CATG1, CATG2') \n";
            //strSQL = strSQL + " ORDER BY WIPINV, CATG1, CATG2') \n";

            return strSQL;
        }

        private string getSQL7()
        {
            string date = tb_to_yyyymmdd.Text.Substring(0, 4) + "-" + tb_to_yyyymmdd.Text.Substring(4, 2) + "-" + tb_to_yyyymmdd.Text.Substring(6, 2);
            /* 실적 조회 쿼리*/
            DateTime bfday = Convert.ToDateTime(date);
            string str = bfday.Year.ToString("0000") + bfday.Month.ToString("00") + bfday.Day.ToString("00");

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("	USP_INCOME_STOCK_SEARCH  '" + str + "' \n");
            return sbSQL.ToString();
        }
    }
}