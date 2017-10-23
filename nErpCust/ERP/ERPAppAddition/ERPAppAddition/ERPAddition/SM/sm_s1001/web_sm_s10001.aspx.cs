using System;
using System.Data;
using System.Text;
//using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Oracle.DataAccess.Client;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
//using System.Data.OleDb;
//using System.Data.OracleClient;
using Microsoft.Reporting.WebForms;
using ERPAppAddition.ERPAddition.SM;

namespace ERPAppAddition.ERPAddition.SM.sm_s1001
{
    public partial class web_sm_s10001 : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.AppSettings["connectionKey"];

        string sql_cust_cd;

        OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);
        OracleConnection conn_if = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_UNIERP"].ConnectionString);
        SqlConnection conn_erp= new SqlConnection(ConfigurationManager.ConnectionStrings["nepes"].ConnectionString);

        SqlCommand cmd_erp = new SqlCommand();
        OracleCommand cmd = new OracleCommand();
        OracleDataReader dr;
        SqlDataReader dr_erp;
        OracleDataAdapter sqlAdapter1;
        DataSet ds = new DataSet();

       // t_if_device_amt ds_device_amt = new t_if_device_amt();
        

        FarPoint.Web.Spread.SpreadCommandEventArgs chk;
        string sql_spread;
        int value,chk_save_yn = 0;
        string userid, db_name;
        protected void Page_Load(object sender, EventArgs e)
        {            

            if (!Page.IsPostBack)
            {               

                sql_cust_cd = "SELECT SYSCODE_NAME cust_cd FROM SYSCODEDATA A WHERE  A.PLANT = 'CCUBEDIGITAL' AND A.SYSTABLE_NAME IN ( 'CUSTOMER') UNION ALL SELECT '%' FROM DUAL  ORDER BY 1 ";
                string sql = "";
                ds_sm_s1001 dt1 = new ds_sm_s1001();

                if (Request.QueryString["userid"] == null || Request.QueryString["userid"] == "")
                    userid = "dev"; //erp에서 실행하지 않았을시 대비용
                else
                    userid = Request.QueryString["userid"];
                Session["User"] = userid;//Request.QueryString["userid"];
                //ReportViewer1.Reset();
                ReportCreator(dt1, sql, ReportViewer1, "rv_sm_s1001.rdlc", "DataSet1");
                FillDropDownList(sql_cust_cd);
                FillRadioButton();

                WebSiteCount();
            }

           // FpSpread_amt.Attributes.Add("onDataChanged", "ProfileSpread");
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = conn_erp.Database.ToString();
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private void FillDropDownList(string sql)
        {
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            try
            {
                // 품목 드랍다운리스트 내용을 보여준다.
                OracleCommand cmd2 = new OracleCommand(sql_cust_cd, conn);

                dr = cmd2.ExecuteReader();
                // ListItem liObject = ddl_cust_cd.Items.FindByValue("SEC");
                if (ddl_cust_cd.Items.Count < 2)
                {
                    ddl_cust_cd.DataSource = dr;
                    ddl_cust_cd.DataValueField = "cust_cd";
                    ddl_cust_cd.DataTextField = "cust_cd";
                    ddl_cust_cd.DataBind();
                }
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void FillRadioButton()
        {
            string sql;
            //사용자 권한중 단가 권한이 있는지를 확인후 있으면 기준정보 입력화면을 셋팅해줌 없으면 집계 A /상세 B 만 ( 기준정보 C)
            sql = "select A.USR_ID  from z_usr_mast_rec_usr_role_asso  a inner join z_usr_mast_rec B ON A.USR_ID = B.USR_ID " +
                  " where usr_role_id like '%SA-PRICE00%' and USE_YN = 'y' and A.USR_ID = '" + Session["User"].ToString() + "' ";
                        
            DataTable dt = Execute_ERP(sql);

            int chk_userid;
            chk_userid = dt.Rows.Count;
            if (chk_userid == 0)
            {
                rbl_view_type.Items.Add((new ListItem("집계", "A")));
                rbl_view_type.Items.Add((new ListItem("상세", "B")));
            }
            else
            {
                rbl_view_type.Items.Add((new ListItem("집계", "A")));
                rbl_view_type.Items.Add((new ListItem("상세", "B")));
                rbl_view_type.Items.Add((new ListItem("기준정보", "C")));
            }
                
            rbl_view_type.SelectedIndex = 0;

        }

        private DataTable Execute_ERP(string sql)
        {
            conn_erp.Open();
            cmd_erp = conn_erp.CreateCommand();
            cmd_erp.CommandType = CommandType.Text;
            cmd_erp.CommandText = sql;
            DataTable dt = new DataTable();
            try
            {
                // 품목 드랍다운리스트 내용을 보여준다.
                //SqlConnection cmd2 = new SqlConnection(conn_erp);
                
                dr_erp = cmd_erp.ExecuteReader();
                dt.Load(dr_erp);                
            }
            catch (Exception ex)
            {
                if (conn_erp.State == ConnectionState.Open)
                    conn_erp.Close();
            }
            conn_erp.Close();
            return dt;
        }
        private void ReportCreator(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {

            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {
                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_" + pg_gubun.SelectedItem + DateTime.Now.ToShortDateString(); 
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                _reportViewer.LocalReport.DataSources.Add(rds);
                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }

        }

        private void ReportCreator2(DataSet _dataSet, string _Query, ReportViewer _reportViewer, string _ReportName, string _ReportDataSourceName)
        {

            conn_if.Open();
            cmd = conn_if.CreateCommand();
            cmd.CommandType = CommandType.Text;

            DataSet ds = _dataSet;
            try
            {               

                cmd.CommandText = _Query;
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                dr.Close();
                _reportViewer.LocalReport.ReportPath = Server.MapPath(_ReportName);

                _reportViewer.LocalReport.DisplayName = "REPORT_" + pg_gubun.SelectedItem + DateTime.Now.ToShortDateString();
                ReportDataSource rds = new ReportDataSource();
                rds.Name = _ReportDataSourceName;
                rds.Value = ds.Tables[0];
                _reportViewer.LocalReport.DataSources.Add(rds);
                _reportViewer.LocalReport.Refresh();
            }
            catch { }
            finally
            {
                if (conn_if.State == ConnectionState.Open)
                    conn_if.Close();
            }

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string fr_dt, to_dt, cust_cd;fr_dt = str_fr_dt.Text.Trim() + "070000";
            to_dt = DateTime.ParseExact(str_to_dt.Text.Substring(0, 8).ToLower(), "yyyyMMdd", null).AddMonths(0).AddDays(1).ToString("yyyyMMdd") + "070000";
            cust_cd = ddl_cust_cd.Text.Trim();
            
            /* 
             * 의뢰자: 박은아 , 개발자: 송태호
             * 출하실적중 동부하이텍 범프수량: 반제품출하 가 마이티상 셋팅이 안되어서 마이티 레포트상 동부하이텍생산실적 수량으로 셋팅 
             * file 수량: f/t (final test)수량중에서 개발수량 제외
             * 
             */

            ds_sm_s1001 dt1 = new ds_sm_s1001();
            ReportViewer1.Reset();

            //집계조회
            if (rbl_view_type.SelectedValue == "A")
            {

                //수주실적집계
                if (pg_gubun.Text == "PG_A1")
                {
                    string sql
                    = "WITH DATA_HOUSE AS ( " +
                        "SELECT DD, CUSTOMER, CUSTOMER_NM, PROC_TYPE,TYPE_GP, TYPE_GP2, WAFER_DIA,IMPORT_QTY_UNIT,SUM(IMPORT_QTY) IN_QTY,ROUTESET " +
                        "FROM ( " +
                        "  SELECT DISTINCT A.CUSTOMER_LOT ,SUBSTR(A.IMPORT_TIME,7,2) DD " +
                        "       , B.CUSTOMER " +
                        "       , (SELECT Substr(SYSCODE_DESC, 1, Instr(SYSCODE_DESC, '%', 1, 1) - 1)  " +
                        "          FROM   SYSCODEDATA   " +
                        "          WHERE  PLANT = 'CCUBEDIGITAL'  " +
                        "            AND SYSTABLE_NAME IN ( 'CUSTOMER' )  " +
                        "            AND SYSCODE_NAME = B.CUSTOMER) CUSTOMER_NM  " +
                        "       , A.PART_ID " +
                        "       , B.PKGTYPE " +
                        "       , A.IMPORT_QTY " +
                        "       , A.IMPORT_QTY_UNIT " +
                        "       , A.INPUT_QTY " +
                        "       , WAFER_QTY " +
                        "       , C.SUB_PLANT_1 TYPE_GP " +
                        "       , C.SUB_PLANT_2 TYPE_GP2 " +
                        "       , B.WAFER_DIA " +
                        "       , A.PROC_TYPE, D.ROUTESET " +
                        "    FROM IMPORTLOT A, PARTSPEC B,PROC_TYPE_INFO C, PARTROUTESET D " +
                        "   WHERE     A.PLANT = 'CCUBEDIGITAL' " +
                        "         AND A.PLANT = B.PLANT " +
                        "         AND A.PLANT = C.PLANT AND  A.PLANT = D.PLANT" +
                        "         AND B.PROC_TYPE = C.PROC_TYPE " +
                        "         AND A.PART_ID = B.PART_ID AND A.PART_ID = D.PART_ID " +
                        "         AND B.CUSTOMER LIKE '" + ddl_cust_cd.Text.Trim() + "' " +
                        "         AND A.IMPORT_TIME between '" + fr_dt.Substring(0, 8) + "' AND  decode('" + fr_dt.Substring(0, 8) + "','" + to_dt.Substring(0, 8) + "','" + to_dt.Substring(0, 8) + "',to_char(to_date('" + to_dt.Substring(0, 8) + "','yyyymmdd')-1,'YYYYMMDD')) " +
                        "         AND C.IN_OPER IN (  SELECT MIN( operation_old) FROM LOTHST z WHERE  PART_NEW = A.PART_ID  AND PLANT = 'CCUBEDIGITAL' and main_lot = A.CUSTOMER_LOT and trans_time like '" + fr_dt.Substring(0, 6) + "'||'%'  and lot_number in ( select lot_number  from lotsts  where PLANT = 'CCUBEDIGITAL' and main_lot = z.main_lot) )" +
                        " )  " +
                        " GROUP BY DD, CUSTOMER,CUSTOMER_NM, PROC_TYPE,TYPE_GP, TYPE_GP2,WAFER_DIA,IMPORT_QTY_UNIT,ROUTESET " +
                        "ORDER BY DD " +
                        ") " +
                        " SELECT A.CUSTOMER, A.DD, A.WAFER_DIA INCH, B.ITEM_GROUP, B.SEQ, SUM(A.IN_QTY) QTY " +
                        "   FROM DATA_HOUSE A ,  (select A4.item_group, B3.routeset, A4.REMARK SEQ " +
                        "                           from UNIERPSEMI.t_device_group A4, UNIERPSEMI.t_device_group_routeset B3 " +
                        "                          where  A4.item_group = B3.item_group " +
                        "                           AND A4.REMARK <> '예외거래처용'  ) B " +
                        "  where a.routeset = b.routeset and b.item_group not in ( 'BG','CHIP BIZ','FILM')  " +  //제외항목추가 20131122 FILM
                        "  GROUP BY A.CUSTOMER, A.DD, A.WAFER_DIA, B.ITEM_GROUP, B.SEQ " +
                        " ORDER BY DD, CUSTOMER, SEQ ";
                    ReportCreator(dt1, sql, ReportViewer1, "rv_sm_s1001_v4.rdlc", "DataSet1");
                }
                //출하실적집계                
                if (pg_gubun.Text == "PG_B1")
                {
                    string frdt, todt;
                    frdt = fr_dt.Substring(0, 8);
                    todt = to_dt.Substring(0, 8);
                    string sql
                = "WITH  DATAHOUSE AS (  SELECT customer,  " +
                    "       dd,  " +
                    "       a.item_group,  " +
                    "       inch,  " +
                    "       SUM(qty) qty, SUM(ASSY_OUT_QTY) ASSY_OUT_QTY, b.remark seq,LOT_REASON  " +
                    "  FROM (SELECT DISTINCT A2.CUSTOMER,  " +
                    "               SUBSTR(DD, 7, 2) DD,  " +
                    "               B2.item_group,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN 0  " +
                    "                    ELSE inch END inch,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN SUM(OUT_QTY)  " +
                    "                    ELSE SUM(IN_QTY) END QTY  , sum(ASSY_OUT_QTY) ASSY_OUT_QTY,LOT_REASON " +
                    "          FROM (SELECT /*+ ORDERED */ SUBSTR(A3.ISSUE_TIME, 1, 8) DD,  " +
                    "                       B1.IN_QTY,  " +
                    "                       B1.IN_QTY_UNIT,  " +
                    "                       B1.OUT_QTY,  " +
                    "                       B1.OUT_QTY_UNIT,  " +
                    "                       CASE WHEN ROUTESET LIKE 'INK-ONLY-T%' THEN ROUTESET  " +
                    "                            ELSE (SELECT ROUTESET  " +
                    "                                    FROM PARTROUTESET@MIGHTY  " +
                    "                                   WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                     AND PART_ID = (SELECT PART  " +
                    "                                                      FROM PACKINGLIST@MIGHTY  " +
                    "                                                     WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                                       AND PACKINGLIST_NO = A3.PACKINGLIST_NO)) END ROUTE,  " +
                    "                       C.CUSTOMER,  " +
                    "                       (SELECT WAFER_DIA  " +
                    "                          FROM PARTSPEC@mighty  " +
                    "                         WHERE part_id = A3.part  " +
                    "                           AND plant = A3.plant) inch,  " +
                    "                       PACKING_TYPE,  " +
                    "                       NVL(B1.LOT_REASON, 'ZZ') LOT_REASON ,DECODE (B1.ASSY_OUT_QTY, 0,  B1.OUT_QTY, B1.ASSY_OUT_QTY) AS ASSY_OUT_QTY " +
                    "                  FROM PACKINGLIST@MIGHTY A3,  " +
                    "                       PACKINGLOT@MIGHTY B1,  " +
                    "                       LOTSTS@MIGHTY C  " +
                    "                 WHERE A3.PLANT = B1.PLANT  " +
                    "                   AND A3.PLANT = C.PLANT  " +
                    "                   AND A3.PLANT = 'CCUBEDIGITAL'  " +
                    "                   AND A3.PACKINGLIST_NO = B1.PACKINGLIST_NO  " +
                    "                   AND B1.LOT_NUMBER = C.LOT_NUMBER  " +
                    "                   AND C.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND SUBSTR(A3.ISSUE_TIME, 1, 8) between  '" + frdt + "' and '" + todt + "'  " +
                    "                 ORDER BY A3.ISSUE_TIME, A3.PACKING_TYPE) A2,  " +
                    "               (SELECT A4.item_group,  " +
                    "                       B3.routeset,  " +
                    "                       d.customer,  " +
                    "                       d.packingtype  " +
                    "                  FROM t_device_group A4,  " +
                    "                       t_device_group_routeset B3,  " +
                    "                       t_device_group_cust_packing d  " +
                    "                 WHERE D.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND A4.item_group = d.item_group  " +
                    "                   AND A4.item_group = B3.item_group) B2  " +
                    "         WHERE (A2.LOT_REASON IN (SELECT 'ZZ'  " +
                    "                                    FROM DUAL)  " +
                    "                 OR A2.LOT_REASON IN (SELECT DISTINCT REASON_CODE  " +
                    "                                        FROM t_device_group_reason_code))  " +
                    "           AND A2.ROUTE = B2.routeset  " +
                    "           AND A2.customer = B2.customer  " +
                    "           AND A2.PACKING_TYPE = B2.packingtype  " +
                    "         GROUP BY A2.CUSTOMER, SUBSTR(DD, 7, 2), B2.item_group, inch,LOT_REASON) A inner join t_device_group b on a.item_group = b.item_group " +
                    " GROUP BY customer, dd, a.item_group, inch ,b.remark,LOT_REASON " +
                        //" ORDER BY dd ";
                    " union all  " +
                    " /*동부하이텍 bump 수량 */ " +
                    " SELECT CUSTOMER, dd, ITEM_GROUP,INCH, sum(A.INPUT_QTY) qty,0, (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT CUSTOMER,LH.QTY_1 AS INPUT_QTY   ,SUBSTR ( REPORT_DATE, 7, 2 ) dd  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(LH.PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = LH.part and plant = LH.plant ) inch  " +
                    "              ,'DDI_BUMP' ITEM_GROUP  " +
                    "         FROM WIPHST@mighty LH   " +
                    "         WHERE PLANT = 'CCUBEDIGITAL'    " +
                    "             AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    "             AND CUSTOMER like 'DONGBU HITEK'    " +
                    "             AND OPERATION = '4020'    " +
                    "             AND OPERATION <> TO_OPERATION  " +
                    "             AND CREATE_CODE NOT IN ('PR','PRLM')    " +
                    "   ) A   " +
                    "   where  owner = 'NL'    " +
                    " GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP  " +
                    " UNION ALL  " +
                    " /*동부하이텍/ 매그나칩 bg수량  */" +
                    " SELECT  CUSTOMER, dd, ITEM_GROUP,INCH,SUM(A.INPUT_QTY) QTY,0 , (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON  " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT B.CUSTOMER,SUBSTR ( REPORT_DATE, 7, 2 ) dd ,a.IN_QTY AS INPUT_QTY  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = B.part and plant = B.plant ) inch   " +
                    "              , 'BG' ITEM_GROUP  " +
                    "           FROM  PACKINGLOT@mighty A, WIPHST@mighty B      " +
                    "          WHERE A.PLANT (+) = B.PLANT     AND B.PLANT  ='CCUBEDIGITAL'      " +
                    "            AND A.LOT_NUMBER (+) = B.LOT_NUMBER     " +
                    "            AND B.CUSTOMER IN ( 'DONGBU HITEK','MAGNACHIP')       " +
                    "            AND OPERATION <> TO_OPERATION      " +
                    "            AND OPERATION ='5800'       " +
                    "            AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    " ) A   " +
                    "   where  owner = 'NL'    " +
                    "   GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP  " +
                    " )  " +
                    " SELECT CUSTOMER, DD, INCH, ITEM_GROUP, SUM(QTY) QTY, SEQ  " +
                    " FROM ( SELECT CUSTOMER, DD, INCH " +
                        //"      , ITEM_GROUP  " +
                    "             , CASE WHEN  CUSTOMER in ('DONGBU HITEK', 'LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_TEST' and LOT_REASON = 'ER-CUST_PLT' THEN 'DDI_BUMP'" + //20131014 박은아요청 이거래처들꺼만  품목그룹이 바뀌어야 한다.
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then 'DDI_BUMP'  " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then 'WLP_BUMP'  " +
                    "                    ELSE ITEM_GROUP END ITEM_GROUP " +
                    "             , CASE WHEN CUSTOMER = 'DONGBU HITEK' AND ITEM_GROUP = 'ASSY' THEN ASSY_OUT_QTY   " +
                    "                    WHEN CUSTOMER = 'LAPIS' AND ITEM_GROUP = 'CHIP BIZ' THEN (SELECT QTY FROM DATAHOUSE WHERE CUSTOMER = 'LAPIS' AND DD = A.DD AND ITEM_GROUP = 'COG' ) " +
                    "                    ELSE  QTY END QTY " +
                    "              ,case when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then '1' " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then '3' " +
                    "                    else SEQ end seq  " +
                    "           FROM DATAHOUSE  A   " +
                    "          WHERE CUSTOMER LIKE  '" + cust_cd + "' ) Z  " +
                    "  GROUP BY CUSTOMER, DD, INCH, ITEM_GROUP,SEQ  ORDER BY DD ";



                    ReportCreator2(dt1, sql, ReportViewer1, "rv_sm_s1001_v4.rdlc", "DataSet1");

                }
                //매출실적집계                
                if (pg_gubun.Text == "PG_C")
                {
                    string frdt, todt;
                    frdt = fr_dt.Substring(0, 8);
                    todt = to_dt.Substring(0, 8);
                    string sql
                = "WITH  DATAHOUSE AS (  SELECT customer,  " +
                    "       dd,  " +
                    "       a.item_group, PART,  " +
                    "       inch,  " +
                    "       SUM(qty) qty, SUM(ASSY_OUT_QTY) ASSY_OUT_QTY, b.remark seq,LOT_REASON  " +
                    "  FROM (SELECT DISTINCT A2.CUSTOMER,  " +
                    "               SUBSTR(DD, 7, 2) DD,  " +
                    "               B2.item_group, PART,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN 0  " +
                    "                    ELSE inch END inch,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN SUM(OUT_QTY)  " +
                    "                    ELSE SUM(IN_QTY) END QTY  , sum(ASSY_OUT_QTY) ASSY_OUT_QTY,LOT_REASON " +
                    "          FROM (SELECT /*+ ORDERED */ SUBSTR(A3.ISSUE_TIME, 1, 8) DD, A3.PART, " +
                    "                       B1.IN_QTY,  " +
                    "                       B1.IN_QTY_UNIT,  " +
                    "                       B1.OUT_QTY,  " +
                    "                       B1.OUT_QTY_UNIT,  " +
                    "                       CASE WHEN ROUTESET LIKE 'INK-ONLY-T%' THEN ROUTESET  " +
                    "                            ELSE (SELECT ROUTESET  " +
                    "                                    FROM PARTROUTESET@MIGHTY  " +
                    "                                   WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                     AND PART_ID = (SELECT PART  " +
                    "                                                      FROM PACKINGLIST@MIGHTY  " +
                    "                                                     WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                                       AND PACKINGLIST_NO = A3.PACKINGLIST_NO)) END ROUTE,  " +
                    "                       C.CUSTOMER,  " +
                    "                       (SELECT WAFER_DIA  " +
                    "                          FROM PARTSPEC@mighty  " +
                    "                         WHERE part_id = A3.part  " +
                    "                           AND plant = A3.plant) inch,  " +
                    "                       PACKING_TYPE,  " +
                    "                       NVL(B1.LOT_REASON, 'ZZ') LOT_REASON ,DECODE (B1.ASSY_OUT_QTY, 0,  B1.OUT_QTY, B1.ASSY_OUT_QTY) AS ASSY_OUT_QTY " +
                    "                  FROM PACKINGLIST@MIGHTY A3,  " +
                    "                       PACKINGLOT@MIGHTY B1,  " +
                    "                       LOTSTS@MIGHTY C  " +
                    "                 WHERE A3.PLANT = B1.PLANT  " +
                    "                   AND A3.PLANT = C.PLANT  " +
                    "                   AND A3.PLANT = 'CCUBEDIGITAL'  " +
                    "                   AND A3.PACKINGLIST_NO = B1.PACKINGLIST_NO  " +
                    "                   AND B1.LOT_NUMBER = C.LOT_NUMBER  " +
                    "                   AND C.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND SUBSTR(A3.ISSUE_TIME, 1, 8) between  '" + frdt + "' and '" + todt + "'  " +
                    "                 ORDER BY A3.ISSUE_TIME, A3.PACKING_TYPE) A2,  " +
                    "               (SELECT A4.item_group,  " +
                    "                       B3.routeset,  " +
                    "                       d.customer,  " +
                    "                       d.packingtype  " +
                    "                  FROM t_device_group A4,  " +
                    "                       t_device_group_routeset B3,  " +
                    "                       t_device_group_cust_packing d  " +
                    "                 WHERE D.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND A4.item_group = d.item_group  " +
                    "                   AND A4.item_group = B3.item_group) B2  " +
                    "         WHERE (A2.LOT_REASON IN (SELECT 'ZZ'  " +
                    "                                    FROM DUAL)  " +
                    "                 OR A2.LOT_REASON IN (SELECT DISTINCT REASON_CODE  " +
                    "                                        FROM t_device_group_reason_code))  " +
                    "           AND A2.ROUTE = B2.routeset  " +
                    "           AND A2.customer = B2.customer  " +
                    "           AND A2.PACKING_TYPE = B2.packingtype  " +
                    "         GROUP BY A2.CUSTOMER, SUBSTR(DD, 7, 2), B2.item_group, PART, inch,LOT_REASON) A inner join t_device_group b on a.item_group = b.item_group " +
                    " GROUP BY customer, dd, a.item_group, PART, inch ,b.remark,LOT_REASON " +
                        //" ORDER BY dd ";
                    " union all  " +
                    " /*동부하이텍 bump 수량 */ " +
                    " SELECT CUSTOMER, dd, ITEM_GROUP, PART,INCH, sum(A.INPUT_QTY) qty,0, (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT CUSTOMER, PART ,LH.QTY_1 AS INPUT_QTY   ,SUBSTR ( REPORT_DATE, 7, 2 ) dd  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(LH.PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = LH.part and plant = LH.plant ) inch  " +
                    "              ,'DDI_BUMP' ITEM_GROUP  " +
                    "         FROM WIPHST@mighty LH   " +
                    "         WHERE PLANT = 'CCUBEDIGITAL'    " +
                    "             AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    "             AND CUSTOMER like 'DONGBU HITEK'    " +
                    "             AND OPERATION = '4020'    " +
                    "             AND OPERATION <> TO_OPERATION  " +
                    "             AND CREATE_CODE NOT IN ('PR','PRLM')    " +
                    "   ) A   " +
                    "   where  owner = 'NL'    " +
                    " GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP , PART " +
                    " UNION ALL  " +
                    " /*동부하이텍/ 매그나칩 bg수량  */" +
                    " SELECT  CUSTOMER, dd, ITEM_GROUP, PART,INCH,SUM(A.INPUT_QTY) QTY,0 , (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON  " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT B.CUSTOMER, B.PART,SUBSTR ( REPORT_DATE, 7, 2 ) dd ,a.IN_QTY AS INPUT_QTY  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = B.part and plant = B.plant ) inch   " +
                    "              , 'BG' ITEM_GROUP  " +
                    "           FROM  PACKINGLOT@mighty A, WIPHST@mighty B      " +
                    "          WHERE A.PLANT (+) = B.PLANT     AND B.PLANT  ='CCUBEDIGITAL'      " +
                    "            AND A.LOT_NUMBER (+) = B.LOT_NUMBER     " +
                    "            AND B.CUSTOMER IN ( 'DONGBU HITEK','MAGNACHIP')       " +
                    "            AND OPERATION <> TO_OPERATION      " +
                    "            AND OPERATION ='5800'       " +
                    "            AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    " ) A   " +
                    "   where  owner = 'NL'    " +
                    "   GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP, PART  " +
                    " )  " +
                    " SELECT CUSTOMER, DD, INCH, ITEM_GROUP, SUM(QTY) QTY, SEQ " +
                    " FROM (" +
                    "    SELECT CUSTOMER, DD, INCH, Z.ITEM_GROUP,Z.PART, SUM(Z.QTY) * SUM(Y.AMT) QTY, SEQ  " +
                    "    FROM ( SELECT CUSTOMER, DD, INCH " +
                    "             , CASE WHEN  CUSTOMER in ('DONGBU HITEK', 'LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_TEST' and LOT_REASON = 'ER-CUST_PLT' THEN 'DDI_BUMP'" + //20131014 박은아요청 이거래처들꺼만  품목그룹이 바뀌어야 한다.
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then 'DDI_BUMP'  " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then 'WLP_BUMP'  " +
                    "                    ELSE ITEM_GROUP END ITEM_GROUP , PART  " +
                    "             , CASE WHEN CUSTOMER = 'DONGBU HITEK' AND ITEM_GROUP = 'ASSY' THEN ASSY_OUT_QTY   " +
                    "                    WHEN CUSTOMER = 'LAPIS' AND ITEM_GROUP = 'CHIP BIZ' THEN (SELECT QTY FROM DATAHOUSE WHERE CUSTOMER = 'LAPIS' AND DD = A.DD AND ITEM_GROUP = 'COG' ) " +
                    "                    ELSE  QTY END QTY " +
                    "              ,case when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then '1' " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then '3' " +
                    "                    else SEQ end seq  " +
                    "           FROM DATAHOUSE  A   " +
                    "          WHERE CUSTOMER LIKE  '" + cust_cd + "' ) Z INNER JOIN t_device_amt Y ON Z.PART = Y.PART AND Z.ITEM_GROUP = Y.ITEM_GROUP " +
                    "    WHERE Y.YYYYMM BETWEEN  '" + frdt.Substring(0, 6) + "' and '" + todt.Substring(0, 6) + "'  " +
                    "    GROUP BY CUSTOMER, DD, INCH, Z.ITEM_GROUP,Z.PART,SEQ , Y.AMT ORDER BY DD " +
                    " ) A GROUP BY CUSTOMER, DD, INCH, ITEM_GROUP , SEQ";



                    ReportCreator2(dt1, sql, ReportViewer1, "rv_sm_s1001_v5.rdlc", "DataSet1");
                }
            }
            else //상세조회에서
            {

                //수주상세
                if (pg_gubun.Text == "PG_A1")
                {
                    string sql
                    = "WITH DATA_HOUSE AS ( " +
                        "SELECT DD, CUSTOMER, CUSTOMER_NM, PART_ID part, PROC_TYPE,TYPE_GP, TYPE_GP2, WAFER_DIA,IMPORT_QTY_UNIT,SUM(IMPORT_QTY) IN_QTY,ROUTESET " +
                        "FROM ( " +
                        "  SELECT DISTINCT A.CUSTOMER_LOT ,SUBSTR(A.IMPORT_TIME,7,2) DD " +
                        "       , B.CUSTOMER " +
                        "       , (SELECT Substr(SYSCODE_DESC, 1, Instr(SYSCODE_DESC, '%', 1, 1) - 1)  " +
                        "          FROM   SYSCODEDATA   " +
                        "          WHERE  PLANT = 'CCUBEDIGITAL'  " +
                        "            AND SYSTABLE_NAME IN ( 'CUSTOMER' )  " +
                        "            AND SYSCODE_NAME = B.CUSTOMER) CUSTOMER_NM  " +
                        "       , A.PART_ID " +
                        "       , B.PKGTYPE " +
                        "       , A.IMPORT_QTY " +
                        "       , A.IMPORT_QTY_UNIT " +
                        "       , A.INPUT_QTY " +
                        "       , WAFER_QTY " +
                        "       , C.SUB_PLANT_1 TYPE_GP " +
                        "       , C.SUB_PLANT_2 TYPE_GP2 " +
                        "       , B.WAFER_DIA " +
                        "       , A.PROC_TYPE, D.ROUTESET " +
                        "    FROM IMPORTLOT A, PARTSPEC B,PROC_TYPE_INFO C, PARTROUTESET D " +
                        "   WHERE     A.PLANT = 'CCUBEDIGITAL' " +
                        "         AND A.PLANT = B.PLANT " +
                        "         AND A.PLANT = C.PLANT AND  A.PLANT = D.PLANT" +
                        "         AND B.PROC_TYPE = C.PROC_TYPE " +
                        "         AND A.PART_ID = B.PART_ID AND A.PART_ID = D.PART_ID " +
                        "         AND B.CUSTOMER LIKE '" + ddl_cust_cd.Text.Trim() + "' " +
                        "         AND A.IMPORT_TIME between '" + fr_dt.Substring(0, 8) + "' AND  decode('" + fr_dt.Substring(0, 8) + "','" + to_dt.Substring(0, 8) + "','" + to_dt.Substring(0, 8) + "',to_char(to_date('" + to_dt.Substring(0, 8) + "','yyyymmdd')-1,'YYYYMMDD')) " +
                        "         AND C.IN_OPER IN (  SELECT MIN( operation_old) FROM LOTHST z WHERE  PART_NEW = A.PART_ID  AND PLANT = 'CCUBEDIGITAL' and main_lot = A.CUSTOMER_LOT and trans_time like '" + fr_dt.Substring(0, 6) + "'||'%'  and lot_number in ( select lot_number  from lotsts  where PLANT = 'CCUBEDIGITAL' and main_lot = z.main_lot) )" +
                        " )  " +
                        " GROUP BY DD, CUSTOMER,CUSTOMER_NM,PART_ID ,PROC_TYPE,TYPE_GP, TYPE_GP2,WAFER_DIA,IMPORT_QTY_UNIT,ROUTESET " +
                        "ORDER BY DD " +
                        ") " +
                        " SELECT A.CUSTOMER, A.DD, part, A.WAFER_DIA INCH, B.ITEM_GROUP, B.SEQ, SUM(A.IN_QTY) QTY " +
                        "   FROM DATA_HOUSE A ,  (select A4.item_group, B3.routeset, A4.REMARK SEQ " +
                        "                           from UNIERPSEMI.t_device_group A4, UNIERPSEMI.t_device_group_routeset B3 " +
                        "                          where  A4.item_group = B3.item_group " +
                        "                           AND A4.REMARK <> '예외거래처용'  ) B " +
                        "  where a.routeset = b.routeset and b.item_group not in ( 'BG','CHIP BIZ','FILM')  " +  //제외항목추가 20131122 FILM
                        "  GROUP BY A.CUSTOMER, A.DD,part, A.WAFER_DIA, B.ITEM_GROUP, B.SEQ " +
                        " ORDER BY DD, CUSTOMER, SEQ ";
                    ReportCreator(dt1, sql, ReportViewer1, "rv_sm_s1001_v4_detail.rdlc", "DataSet1");
                }

                //출하상세
                if (pg_gubun.Text == "PG_B1")
                {
                    string frdt, todt;
                    frdt = fr_dt.Substring(0, 8);
                    todt = to_dt.Substring(0, 8);
                    string sql
                  = "WITH  DATAHOUSE AS (  SELECT customer,  " +
                    "       dd,  " +
                    "       a.item_group, PART,  " +
                    "       inch,  " +
                    "       SUM(qty) qty, SUM(ASSY_OUT_QTY) ASSY_OUT_QTY, b.remark seq,LOT_REASON  " +
                    "  FROM (SELECT DISTINCT A2.CUSTOMER,  " +
                    "               SUBSTR(DD, 7, 2) DD,  " +
                    "               B2.item_group, PART,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN 0  " +
                    "                    ELSE inch END inch,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN SUM(OUT_QTY)  " +
                    "                    ELSE SUM(IN_QTY) END QTY  , sum(ASSY_OUT_QTY) ASSY_OUT_QTY,LOT_REASON " +
                    "          FROM (SELECT /*+ ORDERED */ SUBSTR(A3.ISSUE_TIME, 1, 8) DD, A3.PART, " +
                    "                       B1.IN_QTY,  " +
                    "                       B1.IN_QTY_UNIT,  " +
                    "                       B1.OUT_QTY,  " +
                    "                       B1.OUT_QTY_UNIT,  " +
                    "                       CASE WHEN ROUTESET LIKE 'INK-ONLY-T%' THEN ROUTESET  " +
                    "                            ELSE (SELECT ROUTESET  " +
                    "                                    FROM PARTROUTESET@MIGHTY  " +
                    "                                   WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                     AND PART_ID = (SELECT PART  " +
                    "                                                      FROM PACKINGLIST@MIGHTY  " +
                    "                                                     WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                                       AND PACKINGLIST_NO = A3.PACKINGLIST_NO)) END ROUTE,  " +
                    "                       C.CUSTOMER,  " +
                    "                       (SELECT WAFER_DIA  " +
                    "                          FROM PARTSPEC@mighty  " +
                    "                         WHERE part_id = A3.part  " +
                    "                           AND plant = A3.plant) inch,  " +
                    "                       PACKING_TYPE,  " +
                    "                       NVL(B1.LOT_REASON, 'ZZ') LOT_REASON ,DECODE (B1.ASSY_OUT_QTY, 0,  B1.OUT_QTY, B1.ASSY_OUT_QTY) AS ASSY_OUT_QTY " +
                    "                  FROM PACKINGLIST@MIGHTY A3,  " +
                    "                       PACKINGLOT@MIGHTY B1,  " +
                    "                       LOTSTS@MIGHTY C  " +
                    "                 WHERE A3.PLANT = B1.PLANT  " +
                    "                   AND A3.PLANT = C.PLANT  " +
                    "                   AND A3.PLANT = 'CCUBEDIGITAL'  " +
                    "                   AND A3.PACKINGLIST_NO = B1.PACKINGLIST_NO  " +
                    "                   AND B1.LOT_NUMBER = C.LOT_NUMBER  " +
                    "                   AND C.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND SUBSTR(A3.ISSUE_TIME, 1, 8) between  '" + frdt + "' and '" + todt + "'  " +
                    "                 ORDER BY A3.ISSUE_TIME, A3.PACKING_TYPE) A2,  " +
                    "               (SELECT A4.item_group,  " +
                    "                       B3.routeset,  " +
                    "                       d.customer,  " +
                    "                       d.packingtype  " +
                    "                  FROM t_device_group A4,  " +
                    "                       t_device_group_routeset B3,  " +
                    "                       t_device_group_cust_packing d  " +
                    "                 WHERE D.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND A4.item_group = d.item_group  " +
                    "                   AND A4.item_group = B3.item_group) B2  " +
                    "         WHERE (A2.LOT_REASON IN (SELECT 'ZZ'  " +
                    "                                    FROM DUAL)  " +
                    "                 OR A2.LOT_REASON IN (SELECT DISTINCT REASON_CODE  " +
                    "                                        FROM t_device_group_reason_code))  " +
                    "           AND A2.ROUTE = B2.routeset  " +
                    "           AND A2.customer = B2.customer  " +
                    "           AND A2.PACKING_TYPE = B2.packingtype  " +
                    "         GROUP BY A2.CUSTOMER, SUBSTR(DD, 7, 2), B2.item_group, PART, inch,LOT_REASON) A inner join t_device_group b on a.item_group = b.item_group " +
                    " GROUP BY customer, dd, a.item_group, PART, inch ,b.remark,LOT_REASON " +
                        //" ORDER BY dd ";
                    " union all  " +
                    " /*동부하이텍 bump 수량 */ " +
                    " SELECT CUSTOMER, dd, ITEM_GROUP, PART,INCH, sum(A.INPUT_QTY) qty,0, (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT CUSTOMER, PART ,LH.QTY_1 AS INPUT_QTY   ,SUBSTR ( REPORT_DATE, 7, 2 ) dd  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(LH.PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = LH.part and plant = LH.plant ) inch  " +
                    "              ,'DDI_BUMP' ITEM_GROUP  " +
                    "         FROM WIPHST@mighty LH   " +
                    "         WHERE PLANT = 'CCUBEDIGITAL'    " +
                    "             AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    "             AND CUSTOMER like 'DONGBU HITEK'    " +
                    "             AND OPERATION = '4020'    " +
                    "             AND OPERATION <> TO_OPERATION  " +
                    "             AND CREATE_CODE NOT IN ('PR','PRLM')    " +
                    "   ) A   " +
                    "   where  owner = 'NL'    " +
                    " GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP , PART " +
                    " UNION ALL  " +
                    " /*동부하이텍/ 매그나칩 bg수량  */" +
                    " SELECT  CUSTOMER, dd, ITEM_GROUP, PART,INCH,SUM(A.INPUT_QTY) QTY,0 , (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON  " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT B.CUSTOMER, B.PART,SUBSTR ( REPORT_DATE, 7, 2 ) dd ,a.IN_QTY AS INPUT_QTY  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = B.part and plant = B.plant ) inch   " +
                    "              , 'BG' ITEM_GROUP  " +
                    "           FROM  PACKINGLOT@mighty A, WIPHST@mighty B      " +
                    "          WHERE A.PLANT (+) = B.PLANT     AND B.PLANT  ='CCUBEDIGITAL'      " +
                    "            AND A.LOT_NUMBER (+) = B.LOT_NUMBER     " +
                    "            AND B.CUSTOMER IN ( 'DONGBU HITEK','MAGNACHIP')       " +
                    "            AND OPERATION <> TO_OPERATION      " +
                    "            AND OPERATION ='5800'       " +
                    "            AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    " ) A   " +
                    "   where  owner = 'NL'    " +
                    "   GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP, PART  " +
                    " )  " +
                    " SELECT CUSTOMER, DD, INCH, ITEM_GROUP, PART, SUM(QTY) QTY, SEQ  " +
                    " FROM ( SELECT CUSTOMER, DD, INCH " +
                    "             , CASE WHEN  CUSTOMER in ('DONGBU HITEK', 'LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_TEST' and LOT_REASON = 'ER-CUST_PLT' THEN 'DDI_BUMP'" + //20131014 박은아요청 이거래처들꺼만  품목그룹이 바뀌어야 한다.
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then 'DDI_BUMP'  " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then 'WLP_BUMP'  " +
                    "                    ELSE ITEM_GROUP END ITEM_GROUP , PART  " +
                    "             , CASE WHEN CUSTOMER = 'DONGBU HITEK' AND ITEM_GROUP = 'ASSY' THEN ASSY_OUT_QTY   " +
                    "                    WHEN CUSTOMER = 'LAPIS' AND ITEM_GROUP = 'CHIP BIZ' THEN (SELECT QTY FROM DATAHOUSE WHERE CUSTOMER = 'LAPIS' AND DD = A.DD AND ITEM_GROUP = 'COG' ) " +
                    "                    ELSE  QTY END QTY " +
                    "              ,case when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then '1' " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then '3' " +
                    "                    else SEQ end seq  " +
                    "           FROM DATAHOUSE  A   " +
                    "          WHERE CUSTOMER LIKE  '" + cust_cd + "' ) Z  " +
                    "  GROUP BY CUSTOMER, DD, INCH, ITEM_GROUP, PART,SEQ  ORDER BY DD ";

                    ReportCreator2(dt1, sql, ReportViewer1, "rv_sm_s1001_v4_detail.rdlc", "DataSet1");
                }
                if (pg_gubun.Text == "PG_C")
                {
                    string frdt, todt;
                    frdt = fr_dt.Substring(0, 8);
                    todt = to_dt.Substring(0, 8);
                    string sql
                  = "WITH  DATAHOUSE AS (  SELECT customer,  " +
                    "       dd,  " +
                    "       a.item_group, PART,  " +
                    "       inch,  " +
                    "       SUM(qty) qty, SUM(ASSY_OUT_QTY) ASSY_OUT_QTY, b.remark seq,LOT_REASON  " +
                    "  FROM (SELECT DISTINCT A2.CUSTOMER,  " +
                    "               SUBSTR(DD, 7, 2) DD,  " +
                    "               B2.item_group, PART,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN 0  " +
                    "                    ELSE inch END inch,  " +
                    "               CASE WHEN item_group IN ('SAW', 'COG', 'LM', 'T&R', 'ASSY', 'F-test', 'FVI', 'FILM','TOOL') THEN SUM(OUT_QTY)  " +
                    "                    ELSE SUM(IN_QTY) END QTY  , sum(ASSY_OUT_QTY) ASSY_OUT_QTY,LOT_REASON " +
                    "          FROM (SELECT /*+ ORDERED */ SUBSTR(A3.ISSUE_TIME, 1, 8) DD, A3.PART, " +
                    "                       B1.IN_QTY,  " +
                    "                       B1.IN_QTY_UNIT,  " +
                    "                       B1.OUT_QTY,  " +
                    "                       B1.OUT_QTY_UNIT,  " +
                    "                       CASE WHEN ROUTESET LIKE 'INK-ONLY-T%' THEN ROUTESET  " +
                    "                            ELSE (SELECT ROUTESET  " +
                    "                                    FROM PARTROUTESET@MIGHTY  " +
                    "                                   WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                     AND PART_ID = (SELECT PART  " +
                    "                                                      FROM PACKINGLIST@MIGHTY  " +
                    "                                                     WHERE PLANT = 'CCUBEDIGITAL'  " +
                    "                                                       AND PACKINGLIST_NO = A3.PACKINGLIST_NO)) END ROUTE,  " +
                    "                       C.CUSTOMER,  " +
                    "                       (SELECT WAFER_DIA  " +
                    "                          FROM PARTSPEC@mighty  " +
                    "                         WHERE part_id = A3.part  " +
                    "                           AND plant = A3.plant) inch,  " +
                    "                       PACKING_TYPE,  " +
                    "                       NVL(B1.LOT_REASON, 'ZZ') LOT_REASON ,DECODE (B1.ASSY_OUT_QTY, 0,  B1.OUT_QTY, B1.ASSY_OUT_QTY) AS ASSY_OUT_QTY " +
                    "                  FROM PACKINGLIST@MIGHTY A3,  " +
                    "                       PACKINGLOT@MIGHTY B1,  " +
                    "                       LOTSTS@MIGHTY C  " +
                    "                 WHERE A3.PLANT = B1.PLANT  " +
                    "                   AND A3.PLANT = C.PLANT  " +
                    "                   AND A3.PLANT = 'CCUBEDIGITAL'  " +
                    "                   AND A3.PACKINGLIST_NO = B1.PACKINGLIST_NO  " +
                    "                   AND B1.LOT_NUMBER = C.LOT_NUMBER  " +
                    "                   AND C.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND SUBSTR(A3.ISSUE_TIME, 1, 8) between  '" + frdt + "' and '" + todt + "'  " +
                    "                 ORDER BY A3.ISSUE_TIME, A3.PACKING_TYPE) A2,  " +
                    "               (SELECT A4.item_group,  " +
                    "                       B3.routeset,  " +
                    "                       d.customer,  " +
                    "                       d.packingtype  " +
                    "                  FROM t_device_group A4,  " +
                    "                       t_device_group_routeset B3,  " +
                    "                       t_device_group_cust_packing d  " +
                    "                 WHERE D.CUSTOMER LIKE '" + cust_cd + "'  " +
                    "                   AND A4.item_group = d.item_group  " +
                    "                   AND A4.item_group = B3.item_group) B2  " +
                    "         WHERE (A2.LOT_REASON IN (SELECT 'ZZ'  " +
                    "                                    FROM DUAL)  " +
                    "                 OR A2.LOT_REASON IN (SELECT DISTINCT REASON_CODE  " +
                    "                                        FROM t_device_group_reason_code))  " +
                    "           AND A2.ROUTE = B2.routeset  " +
                    "           AND A2.customer = B2.customer  " +
                    "           AND A2.PACKING_TYPE = B2.packingtype  " +
                    "         GROUP BY A2.CUSTOMER, SUBSTR(DD, 7, 2), B2.item_group, PART, inch,LOT_REASON) A inner join t_device_group b on a.item_group = b.item_group " +
                    " GROUP BY customer, dd, a.item_group, PART, inch ,b.remark,LOT_REASON " +
                        //" ORDER BY dd ";
                    " union all  " +
                    " /*동부하이텍 bump 수량 */ " +
                    " SELECT CUSTOMER, dd, ITEM_GROUP, PART,INCH, sum(A.INPUT_QTY) qty,0, (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT CUSTOMER, PART ,LH.QTY_1 AS INPUT_QTY   ,SUBSTR ( REPORT_DATE, 7, 2 ) dd  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(LH.PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = LH.part and plant = LH.plant ) inch  " +
                    "              ,'DDI_BUMP' ITEM_GROUP  " +
                    "         FROM WIPHST@mighty LH   " +
                    "         WHERE PLANT = 'CCUBEDIGITAL'    " +
                    "             AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    "             AND CUSTOMER like 'DONGBU HITEK'    " +
                    "             AND OPERATION = '4020'    " +
                    "             AND OPERATION <> TO_OPERATION  " +
                    "             AND CREATE_CODE NOT IN ('PR','PRLM')    " +
                    "   ) A   " +
                    "   where  owner = 'NL'    " +
                    " GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP , PART " +
                    " UNION ALL  " +
                    " /*동부하이텍/ 매그나칩 bg수량  */" +
                    " SELECT  CUSTOMER, dd, ITEM_GROUP, PART,INCH,SUM(A.INPUT_QTY) QTY,0 , (SELECT REMARK FROM t_device_group WHERE ITEM_GROUP = A.ITEM_GROUP) SEQ ,'ZZ' LOT_REASON  " +
                    "   FROM   " +
                    "   (   " +
                    "         SELECT B.CUSTOMER, B.PART,SUBSTR ( REPORT_DATE, 7, 2 ) dd ,a.IN_QTY AS INPUT_QTY  " +
                    "              , (CASE WHEN OWNER IS NULL THEN DECODE(PART_TYPE , 'P' , 'NL' ,'ER' )  ELSE OWNER END) OWNER    " +
                    "              , ( select WAFER_DIA from PARTSPEC@mighty where part_id = B.part and plant = B.plant ) inch   " +
                    "              , 'BG' ITEM_GROUP  " +
                    "           FROM  PACKINGLOT@mighty A, WIPHST@mighty B      " +
                    "          WHERE A.PLANT (+) = B.PLANT     AND B.PLANT  ='CCUBEDIGITAL'      " +
                    "            AND A.LOT_NUMBER (+) = B.LOT_NUMBER     " +
                    "            AND B.CUSTOMER IN ( 'DONGBU HITEK','MAGNACHIP')       " +
                    "            AND OPERATION <> TO_OPERATION      " +
                    "            AND OPERATION ='5800'       " +
                    "            AND REPORT_DATE between  '" + frdt + "' and '" + todt + "'  " +
                    " ) A   " +
                    "   where  owner = 'NL'    " +
                    "   GROUP BY   CUSTOMER,dd,INCH,ITEM_GROUP, PART  " +
                    " )  " +
                    " SELECT CUSTOMER, z.DD, INCH, Z.ITEM_GROUP, Z.PART " +
                    "      , DECODE(Y.CURRENCY, 'KRW',SUM(Z.QTY ) * Y.AMT,SUM( Z.QTY ) * Y.AMT " +
                    "        * (SELECT STD_RATE FROM T_IF_DAILY_EXCHANGE_RATE A WHERE YYYYMM = '" + frdt.Substring(0, 6) + "' " +  //일자별 최종 이력 변동환율가져오기
                    "                                      and create_type <> 'C' " + 
                    "                                     AND MES_RECEIVE_DT = (SELECT MAX(MES_RECEIVE_DT) FROM T_IF_DAILY_EXCHANGE_RATE  " +
                    "                                                            WHERE YYYYMM = A.YYYYMM AND CREATE_TYPE = A.CREATE_TYPE AND DD = A.DD AND FROM_CURRENCY = A.FROM_CURRENCY  ) " +
                    "                                     AND DD = Z.DD AND FROM_CURRENCY = Y.CURRENCY) ) QTY  " + 
                    "      , SEQ  " +
                    " FROM ( SELECT CUSTOMER, DD, INCH " +
                    "             , CASE WHEN  CUSTOMER in ('DONGBU HITEK', 'LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_TEST' and LOT_REASON = 'ER-CUST_PLT' THEN 'DDI_BUMP'" + //20131014 박은아요청 이거래처들꺼만  품목그룹이 바뀌어야 한다.
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then 'DDI_BUMP'  " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then 'WLP_BUMP'  " +
                    "                    ELSE ITEM_GROUP END ITEM_GROUP , PART  " +
                    "             , CASE WHEN CUSTOMER = 'DONGBU HITEK' AND ITEM_GROUP = 'ASSY' THEN ASSY_OUT_QTY   " +
                    "                    WHEN CUSTOMER = 'LAPIS' AND ITEM_GROUP = 'CHIP BIZ' THEN (SELECT QTY FROM DATAHOUSE WHERE CUSTOMER = 'LAPIS' AND DD = A.DD AND ITEM_GROUP = 'COG' ) " +
                    "                    ELSE  QTY END QTY " +
                    "              ,case when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'DDI_BUMP ONLY' then '1' " +
                    "                    when CUSTOMER in ('LGE', 'MAGNACHIP', 'TLSI') and ITEM_GROUP = 'WLP_BUMP ONLY' then '3' " +
                    "                    else SEQ end seq  " +
                    "           FROM DATAHOUSE  A   " +
                    "          WHERE CUSTOMER LIKE  '" + cust_cd + "' ) Z INNER JOIN   t_device_amt Y on Z.PART = Y.PART and Z.ITEM_GROUP = Y.ITEM_GROUP " + 
                    "  WHERE Y.YYYYMM =  '" + frdt.Substring(0, 6) + "'  " +
                    "  GROUP BY CUSTOMER, z.DD, INCH, Z.ITEM_GROUP, Z.PART,SEQ , Y.AMT, Y.CURRENCY ORDER BY z.DD ";

                    ReportCreator2(dt1, sql, ReportViewer1, "rv_sm_s1001_v4_detail.rdlc", "DataSet1");
                }
            }

        }

        protected void rbl_view_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbl_view_type.SelectedValue == "C")
            {
                Panel_bas_info.Visible = true;
                rbl_bas_type.Visible = true;
                Panel_Spread_Btn.Visible = true;
                Panel_Spread_bas.Visible = true;
                Panel_Default_Btn.Visible = false;
                Panel_routeset.Visible = false; //라우트셋연결패널
                Panel_default.Visible = false; //레포트뷰어포함된 패널
                panel_device_amt.Visible = false;
            }
            else
            {
                Panel_bas_info.Visible = false;
                rbl_bas_type.Visible = false;
                Panel_Spread_Btn.Visible = false;
                Panel_Spread_bas.Visible = false;
                Panel_Default_Btn.Visible = true;
                Panel_routeset.Visible = false; //라우트셋연결패널
                Panel_default.Visible = true;//레포트뷰어포함된 패널
                panel_device_amt.Visible = false;
            }
            ReportViewer1.Reset();
        }

        protected void btn_Add_Click(object sender, EventArgs e)
        {
            // 품목그룹등록을 선택했으면
            if (rbl_bas_type.SelectedValue == "A")
            {
                if (tb_rowcnt.Text == null || tb_rowcnt.Text == "")
                {
                    MessageBox.ShowMessage("추가할 Row수를 입력해주세요.", this.Page);
                    tb_rowcnt.Focus();
                    return;
                }
                else
                {
                    FpSpread1_ITEMGR.Sheets[0].AddRows(FpSpread1_ITEMGR.Sheets[0].RowCount, Convert.ToInt16(tb_rowcnt.Text));

                }
            }
        }

        protected void btn_save_Click(object sender, EventArgs e)
        {
            // 품목그룹등록을 선택했으면
            if (rbl_bas_type.SelectedValue == "A")
            {
                FpSpread1_ITEMGR.SaveChanges();
                MessageBox.ShowMessage("저장되었습니다.", this.Page);
            }
        }

        protected void FpSpread1_ITEMGR_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            int colcnt;
            int i;                        
            int r = (int)e.CommandArgument;
            colcnt = e.EditValues.Count - 1;



            for (i = 0; i <= colcnt; i++)
            {
                if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
                {
                    string sql;
                    
                    //업데이트시
                    if ( FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value != null)
                    {
                        /*기존값 가져오기*/
                        string ITEM_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[r, 0].Value.ToString();
                        string REMARK = FpSpread1_ITEMGR.Sheets[0].Cells[r, 1].Value.ToString();
                        string cg_ITEM_GROUP, cg_REMARK;
                        /*변경된값 가져오기*/
                        if (i == 0)
                            cg_ITEM_GROUP = e.EditValues[0].ToString();
                        else
                            cg_ITEM_GROUP = ITEM_GROUP;
                        if (i == 1)
                            cg_REMARK = e.EditValues[1].ToString();
                        else
                            cg_REMARK = REMARK;

                        if (cg_REMARK == "System.Object" || cg_REMARK=="")
                            cg_REMARK = " ";
                        sql = "update T_DEVICE_GROUP ";
                        sql = sql + "set ITEM_GROUP = '" + cg_ITEM_GROUP + "',REMARK = '" + cg_REMARK + "',updt_dt = sysdate " ;
                        sql = sql + " where ITEM_GROUP = '" + ITEM_GROUP + "'  ";
                        QueryExecute(conn_if, sql, "");
                    }
                    else
                    {
                        //r = r + 1;
                        //int j = FpSpread1.Sheets[0].ColumnCount;
                        string ITEM_GROUP = e.EditValues[0].ToString();//FpSpread1.Sheets[0].Cells[r, 0].Text;
                        string REMARK = e.EditValues[1].ToString();

                        if (REMARK == "System.Object" || REMARK == "")
                            REMARK = " ";
                        if (ITEM_GROUP == null || ITEM_GROUP == "")
                            MessageBox.ShowMessage("품목그룹명 입력해주세요.", this.Page);                        
                        else
                        {
                            sql = "insert into T_DEVICE_GROUP ";
                            sql = sql + "values('" + ITEM_GROUP + "','" + REMARK + "', 'unierp', sysdate, 'unierp', sysdate)";
                            QueryExecute(conn_if, sql, "");

                        }
                    }
                }
            }
            
        }
        public int QueryExecute(OracleConnection connection,string sql, string wk_type)
        {
            
            connection.Open();
            cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                
                if (wk_type == "check")
                    value = Convert.ToInt32(cmd.ExecuteScalar());
                else
                    value = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                value = -1;
            }

            connection.Close();
            return value;
        }

        public DataTable QueryExeuteDT(OracleConnection connection, string sql)
        {
            ds_sm_s1001_temp ds = new ds_sm_s1001_temp();

            connection.Open();
            cmd = connection.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;

            try
            {
                dr = cmd.ExecuteReader();
                ds.Tables[0].Load(dr);
                connection.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open)
                    connection.Close();
            }

            return ds.Tables[0];
        }
        protected void btn_exe_Click(object sender, EventArgs e)
        {
            string sql;
            // 품목그룹 조회시
            if (rbl_bas_type.SelectedValue == "A")
            {
                sql = "select ITEM_GROUP, REMARK from T_DEVICE_GROUP ";

                sqlAdapter1 = new OracleDataAdapter(sql, conn_if);

                sqlAdapter1.Fill(ds, "ds");

                FpSpread1_ITEMGR.DataSource = ds;
                FpSpread1_ITEMGR.DataBind();
            }
        }

        protected void btn_Delete_Click(object sender, EventArgs e)
        {
            if (rbl_bas_type.SelectedValue == "A")
            {
                System.Collections.IEnumerator enu = FpSpread1_ITEMGR.ActiveSheetView.SelectionModel.GetEnumerator();
                FarPoint.Web.Spread.Model.CellRange cr;

                while (enu.MoveNext())
                {
                    cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                    int a = FpSpread1_ITEMGR.Sheets[0].ActiveRow;
                    //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                    for (int i = 0; i < cr.RowCount; i++)
                    {
                        string ITEM_GROUP = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 0].Text;
                        string REMARK = FpSpread1_ITEMGR.Sheets[0].Cells[FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1].Text;                        


                        string sql = "delete T_DEVICE_GROUP ";
                        sql = sql + " where ITEM_GROUP  ='" + ITEM_GROUP + "' ";

                        if (QueryExecute(conn_if, sql, "") > 0)
                            FpSpread1_ITEMGR.Sheets[0].Rows.Remove(FpSpread1_ITEMGR.Sheets[0].ActiveRow, 1);                        
                    }
                }
            }            

            MessageBox.ShowMessage("삭제되었습니다.", this.Page);
        }
        protected void rbl_bas_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 기본조회용 화면은 감추기
            Panel_default.Visible = false;
            
            if (rbl_bas_type.SelectedValue == "A")
            {
                Panel_Spread_bas.Visible = true;
                Panel_routeset.Visible = false;
                Panel_reasoncode.Visible = false;
                Panel_packingtype.Visible = false;
                panel_e_upload.Visible = false;
                panel_device_amt.Visible = false;
            }
            if (rbl_bas_type.SelectedValue == "B")
            {
                Panel_Spread_bas.Visible = false;
                Panel_routeset.Visible = true;
                Panel_reasoncode.Visible = false;
                Panel_packingtype.Visible = false;
                panel_e_upload.Visible = false;
                panel_device_amt.Visible = false;
            }
            if (rbl_bas_type.SelectedValue == "C")
            {
                Panel_Spread_bas.Visible = false;
                Panel_routeset.Visible = false;
                Panel_reasoncode.Visible = true;
                Panel_packingtype.Visible = false;
                panel_e_upload.Visible = false;
                panel_device_amt.Visible = false;
            }
            // 20130617 packing type추가
            if (rbl_bas_type.SelectedValue == "D")
            {
                Panel_Spread_bas.Visible = false;
                Panel_routeset.Visible = false;
                Panel_reasoncode.Visible = false;
                Panel_packingtype.Visible = true;
                panel_e_upload.Visible = false;
                panel_device_amt.Visible = false;
            }
            // 20131023 매출단가등록
            if (rbl_bas_type.SelectedValue == "E")
            {
                Panel_Spread_bas.Visible = false;
                Panel_routeset.Visible = false;
                Panel_reasoncode.Visible = false;
                Panel_packingtype.Visible = false;
                panel_e_upload.Visible = true;
                panel_device_amt.Visible = false;
            }
            // 20131023 매출단가조회
            if (rbl_bas_type.SelectedValue == "F")
            {
                Panel_Spread_bas.Visible = false;
                Panel_routeset.Visible = false;
                Panel_reasoncode.Visible = false;
                Panel_packingtype.Visible = false;
                panel_e_upload.Visible = false;
                panel_device_amt.Visible = true;
            }
        }
        //***************************품목군과 라우트셋연결 화면 - 조회버튼******************************
        protected void btn_exe_itemgp_routeset_Click(object sender, EventArgs e)
        {
            string sql;
            lsb_l_routeset.Items.Clear(); //내용지우기
            lsb_r_routeset.Items.Clear(); //내용지우기

            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                //왼쪽 라우트셋 가져오기
                sql = "select DISTINCT ROUTESET " +
                      "  from PARTROUTESET@MIGHTY WHERE PLANT = 'CCUBEDIGITAL' " +
                      "   AND PART_ID IN (SELECT PART_ID from PARTSPEC@MIGHTY  A    " +
                      "                    where PLANT = 'CCUBEDIGITAL' AND PART_ID LIKE '%' AND USE_FLAG = 'Y' /*사용여부*/  ) " +
                      "   AND ROUTESET NOT IN (SELECT ROUTESET FROM T_DEVICE_GROUP_ROUTESET where ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "' )  order by 1 ";
                DataTable dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_l_routeset.DataSource = dt;
                    lsb_l_routeset.DataTextField = "ROUTESET";
                    lsb_l_routeset.DataValueField = "ROUTESET";
                    lsb_l_routeset.DataBind();
                }
                //오른쪽 라우트셋 가져오기
                sql = "select ROUTESET from T_DEVICE_GROUP_ROUTESET WHERE ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "'  ";
                dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_r_routeset.DataSource = dt;
                    lsb_r_routeset.DataTextField = "ROUTESET";
                    lsb_r_routeset.DataValueField = "ROUTESET";
                    lsb_r_routeset.DataBind();
                }
            }
        }

        protected void btn_move_right_Click(object sender, EventArgs e)
        {

            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_l_routeset.Items.Count; i++)
                {
                    if (this.lsb_l_routeset.Items[i].Selected)
                    {
                        this.lsb_r_routeset.Items.Add(this.lsb_l_routeset.Items[i]);

                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "insert into T_DEVICE_GROUP_ROUTESET " +
                                     "values ('" + ddl_itemgp.SelectedValue + "', '" + this.lsb_l_routeset.Items[i] + "','unierp', sysdate, 'unierp', sysdate ) ";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);
                        this.lsb_l_routeset.Items.Remove(this.lsb_l_routeset.Items[i]);
                        i--;
                    }
                }

                btn_exe_itemgp_routeset_Click(null, null);
            }
        }

        protected void btn_move_left_Click(object sender, EventArgs e)
        {
            if (ddl_itemgp.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_r_routeset.Items.Count; i++)
                {
                    if (this.lsb_r_routeset.Items[i].Selected)
                    {
                        this.lsb_l_routeset.Items.Add(this.lsb_r_routeset.Items[i]);
                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "delete T_DEVICE_GROUP_ROUTESET " +
                                     "where ITEM_GROUP = '" + ddl_itemgp.SelectedValue + "' and routeset =  '" + this.lsb_r_routeset.Items[i] + "'";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);

                        this.lsb_r_routeset.Items.Remove(this.lsb_r_routeset.Items[i]);
                        i--;

                    }
                }
                btn_exe_itemgp_routeset_Click(null, null);
            }
        }
        //***************************품목군과 이유코드연결 화면 - 조회버튼******************************
        protected void btn_exe_itemgp_reasencode_Click(object sender, EventArgs e)
        {
            string sql;
            lsb_l_reasencode.Items.Clear(); //내용지우기
            lsb_r_reasencode.Items.Clear(); //내용지우기
            if (ddl_itemgp_reasoncode.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_reasoncode.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                //왼쪽 라우트셋 가져오기
                sql = "select SYSCODE_NAME ||' :: ' || SYSCODE_DESC REASON_CODE from SYSCODEDATA@MIGHTY where PLANT = 'CCUBEDIGITAL'  and SYSTABLE_NAME = 'SCRAP_REASON' " +
                      " AND SYSCODE_NAME NOT IN (SELECT REASON_CODE FROM T_DEVICE_GROUP_REASON_CODE where ITEM_GROUP = '" + ddl_itemgp_reasoncode.SelectedValue + "') ORDER BY 1 ";
                DataTable dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_l_reasencode.DataSource = dt;
                    lsb_l_reasencode.DataTextField = "REASON_CODE";
                    lsb_l_reasencode.DataValueField = "REASON_CODE";
                    lsb_l_reasencode.DataBind();
                }
                //오른쪽 라우트셋 가져오기
                sql = "SELECT REASON_CODE ||' :: ' || SYSCODE_DESC REASON_CODE FROM T_DEVICE_GROUP_REASON_CODE a inner join  SYSCODEDATA@MIGHTY b on a.REASON_CODE =b.SYSCODE_NAME and b.PLANT = 'CCUBEDIGITAL'  and b.SYSTABLE_NAME = 'SCRAP_REASON'  WHERE ITEM_GROUP = '" + ddl_itemgp_reasoncode.SelectedValue + "'  ";
                dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_r_reasencode.DataSource = dt;
                    lsb_r_reasencode.DataTextField = "REASON_CODE";
                    lsb_r_reasencode.DataValueField = "REASON_CODE";
                    lsb_r_reasencode.DataBind();
                }
            }
        }

        protected void btn_move_reason_right_Click(object sender, EventArgs e)
        {
            if (ddl_itemgp_reasoncode.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_reasoncode.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_l_reasencode.Items.Count; i++)
                {
                    if (this.lsb_l_reasencode.Items[i].Selected)
                    {
                        this.lsb_r_reasencode.Items.Add(this.lsb_l_reasencode.Items[i]);

                        string reasoncode = this.lsb_l_reasencode.Items[i].ToString();
                        string[] arr = System.Text.RegularExpressions.Regex.Split(reasoncode, " :: ");

                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "insert into T_DEVICE_GROUP_REASON_CODE " +
                                     "values ('" + ddl_itemgp_reasoncode.SelectedValue + "', '" + arr[0] + "','unierp', sysdate, 'unierp', sysdate ) ";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);
                        this.lsb_l_reasencode.Items.Remove(this.lsb_l_reasencode.Items[i]);
                        i--;
                    }
                }

                btn_exe_itemgp_reasencode_Click(null, null);
            }
        }

        protected void btn_move_reason_left_Click(object sender, EventArgs e)
        {
            if (ddl_itemgp_reasoncode.SelectedValue.ToString() == "-선택안됨-" || ddl_itemgp_reasoncode.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_r_reasencode.Items.Count; i++)
                {
                    if (this.lsb_r_reasencode.Items[i].Selected)
                    {
                        this.lsb_l_reasencode.Items.Add(this.lsb_r_reasencode.Items[i]);
                        string reasoncode = this.lsb_r_reasencode.Items[i].ToString();
                         
                        string[] arr = System.Text.RegularExpressions.Regex.Split(reasoncode, " :: ");
                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "delete T_DEVICE_GROUP_REASON_CODE " +
                                     "where ITEM_GROUP = '" + ddl_itemgp_reasoncode.SelectedValue + "' and REASON_CODE =  '" + arr[0] + "'";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);

                        this.lsb_r_reasencode.Items.Remove(this.lsb_r_reasencode.Items[i]);
                        i--;

                    }
                }
                btn_exe_itemgp_reasencode_Click(null, null);
            }
        }
        /*품목군과 거래처별 packing type 연결 */
        /*조회*/
        protected void btn_exe_packingtype_Click(object sender, EventArgs e)
        {
            string sql;
            lsb_l_packingtype.Items.Clear(); //내용지우기
            lsb_r_packingtype.Items.Clear(); //내용지우기
            if (ddl_packingtype_item_gp.SelectedValue.ToString() == "-선택안됨-" || ddl_packingtype_item_gp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else if (ddl_cust_cd2.SelectedValue.ToString() == "-선택안됨-" || ddl_cust_cd2.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("거래처를 선택해 주세요.", this.Page);
            }
            else
            {
                //왼쪽 라우트셋 가져오기
                sql = "select SYSCODE_NAME ||' :: ' || SYSCODE_DESC PACKINGTYPE from SYSCODEDATA@MIGHTY where PLANT = 'CCUBEDIGITAL'  and SYSTABLE_NAME = 'PL_TYPE' " +
                      " AND SYSCODE_NAME NOT IN (SELECT PACKINGTYPE FROM t_device_group_cust_packing where ITEM_GROUP = '" + ddl_packingtype_item_gp.SelectedValue + "' AND CUSTOMER = '" + ddl_cust_cd2.SelectedValue + "' ) ORDER BY 1 ";
                DataTable dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_l_packingtype.DataSource = dt;
                    lsb_l_packingtype.DataTextField = "PACKINGTYPE";
                    lsb_l_packingtype.DataValueField = "PACKINGTYPE";
                    lsb_l_packingtype.DataBind();
                }
                //오른쪽 라우트셋 가져오기
                sql = "SELECT PACKINGTYPE ||' :: ' || SYSCODE_DESC PACKINGTYPE FROM t_device_group_cust_packing a inner join  SYSCODEDATA@MIGHTY b on a.PACKINGTYPE =b.SYSCODE_NAME and b.PLANT = 'CCUBEDIGITAL'  and b.SYSTABLE_NAME = 'PL_TYPE'  WHERE ITEM_GROUP = '" + ddl_packingtype_item_gp.SelectedValue + "'  AND CUSTOMER = '" + ddl_cust_cd2.SelectedValue + "'  ";
                dt = QueryExeuteDT(conn_if, sql);
                if (dt.Rows.Count > 0)
                {
                    lsb_r_packingtype.DataSource = dt;
                    lsb_r_packingtype.DataTextField = "PACKINGTYPE";
                    lsb_r_packingtype.DataValueField = "PACKINGTYPE";
                    lsb_r_packingtype.DataBind();
                }
            }
        }

        protected void btn_move_packingtype_right_Click(object sender, EventArgs e)
        {
            if (ddl_packingtype_item_gp.SelectedValue.ToString() == "-선택안됨-" || ddl_packingtype_item_gp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else if (ddl_cust_cd2.SelectedValue.ToString() == "-선택안됨-" || ddl_cust_cd2.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("거래처를 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_l_packingtype.Items.Count; i++)
                {
                    if (this.lsb_l_packingtype.Items[i].Selected)
                    {
                        this.lsb_r_packingtype.Items.Add(this.lsb_l_packingtype.Items[i]);

                        string packingtype = this.lsb_l_packingtype.Items[i].ToString();
                        string[] arr = System.Text.RegularExpressions.Regex.Split(packingtype, " :: ");

                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "insert into t_device_group_cust_packing " +
                                     "values ('" + ddl_packingtype_item_gp.SelectedValue + "','" + ddl_cust_cd2.SelectedValue + "', '" + arr[0] + "','','unierp', sysdate, 'unierp', sysdate ) ";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);
                        else
                            this.lsb_l_packingtype.Items.Remove(this.lsb_l_packingtype.Items[i]);
                        i--;
                    }
                }

                btn_exe_packingtype_Click(null, null);
            }
        }

        protected void btn_move_packingtype_left_Click(object sender, EventArgs e)
        {
            if (ddl_packingtype_item_gp.SelectedValue.ToString() == "-선택안됨-" || ddl_packingtype_item_gp.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("품목그룹을 선택해 주세요.", this.Page);
            }
            else if (ddl_cust_cd2.SelectedValue.ToString() == "-선택안됨-" || ddl_cust_cd2.SelectedValue.ToString() == null)
            {
                MessageBox.ShowMessage("거래처를 선택해 주세요.", this.Page);
            }
            else
            {
                for (int i = 0; i < lsb_r_packingtype.Items.Count; i++)
                {
                    if (this.lsb_r_packingtype.Items[i].Selected)
                    {
                        this.lsb_l_packingtype.Items.Add(this.lsb_r_packingtype.Items[i]);
                        string packingtype = this.lsb_r_packingtype.Items[i].ToString();

                        string[] arr = System.Text.RegularExpressions.Regex.Split(packingtype, " :: ");
                        //선택된 품목그룹에 왼쪽 리스트박스내용을 insert한다. 
                        string sql = "delete t_device_group_cust_packing " +
                                     "where ITEM_GROUP = '" + ddl_packingtype_item_gp.SelectedValue + "' and CUSTOMER = '" + ddl_cust_cd2.SelectedValue + "' and PACKINGTYPE =  '" + arr[0] + "'";

                        if (QueryExecute(conn_if, sql, "") <= 0)
                            MessageBox.ShowMessage("데이타 저장에 실패했습니다.", this.Page);

                        this.lsb_r_packingtype.Items.Remove(this.lsb_r_packingtype.Items[i]);
                        i--;

                    }
                }
                btn_exe_packingtype_Click(null, null);
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {

            if (FileUpload1.HasFile)
            {
                string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName).ToUpper();
                string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                string FilePath = Server.MapPath(FolderPath + FileName);
                FileUpload1.SaveAs(FilePath);
                if (Extension == ".XLS" || Extension == ".XLSX")
                    GetExcelSheets(FilePath, Extension, "Yes");
                else
                    MessageBox.ShowMessage("Excel 파일만 업로드 가능합니다", this);
            }
        }

        // 엑셀sheet 받아오기
        private void GetExcelSheets(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".XLS": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".XLSX": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                             .ConnectionString;
                    break;
            }

            //Get the Sheets in Excel WorkBoo
            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;
            connExcel.Open();

            //Bind the Sheets to DropDownList
            ddlSheets.Items.Clear();
            ddlSheets.Items.Add(new ListItem("--Select Sheet--", ""));
            ddlSheets.DataSource = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            ddlSheets.DataTextField = "TABLE_NAME";
            ddlSheets.DataValueField = "TABLE_NAME";
            ddlSheets.DataBind();

            DataTable dtCSV = new DataTable();
            OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + ddlSheets.SelectedValue + "]", connExcel);
            OleDbDataAdapter daCSV = new OleDbDataAdapter();
            // OleDbCommand DBCommand;
            IDataReader DBReader;
            DBReader = cmdSelect.ExecuteReader();
            GridView1.DataSource = DBReader;
            GridView1.DataBind();
            
            DBReader.Close();            
            connExcel.Close();            
            HiddenField_fileName.Value = Path.GetFileName(FilePath); //파일명 저장용
            HiddenField_filePath.Value = FilePath; //파일경로 저장용
            HiddenField_extension.Value = Extension; //파일확장자 저장용
            GridView1.Visible = true; //그리드뷰 보여주기
        }

        // 엑셀sheet 선택시 보여주기 위한 함수
        private void ViewExcelSheets(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".XLS": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".XLSX": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                             .ConnectionString;
                    break;
            }

            //Get the Sheets in Excel WorkBoo
            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;
            connExcel.Open();            

            DataTable dtCSV = new DataTable();
            OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + ddlSheets.SelectedValue + "]", connExcel);
            OleDbDataAdapter daCSV = new OleDbDataAdapter();
            // OleDbCommand DBCommand;
            IDataReader DBReader;
            DBReader = cmdSelect.ExecuteReader();
            GridView1.DataSource = DBReader;
            GridView1.DataBind();

            DBReader.Close();
            connExcel.Close();
           
        }
        // 엑셀데이타 저장용 버튼
        protected void btnSave_Click(object sender, EventArgs e)
        {
            
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                string yyyymm = GridView1.Rows[i].Cells[0].Text.Trim();
                string item_group = GridView1.Rows[i].Cells[1].Text.Trim();
                string part = GridView1.Rows[i].Cells[2].Text.Trim();
                string amt = GridView1.Rows[i].Cells[3].Text.Trim();
                string currency  = GridView1.Rows[i].Cells[4].Text.Trim();

                string sql = "select count(part) from t_device_amt where yyyymm = '" + yyyymm + "' and item_group = '" + item_group + "' and part = '" + part + "' ";

                if (QueryExecute(conn_if, sql, "check") > 0) //기존자료가 있으면 삭제
                {
                    //삭제
                    sql = " delete t_device_amt where yyyymm = '" + yyyymm + "' and item_group = '" + item_group + "' and part = '" + part + "' ";
                    if (QueryExecute(conn_if, sql, "") > 0)
                    {
                        //정상 삭제 후 다시 저장
                        sql = " insert into t_device_amt values( '" + yyyymm + "' , '" + item_group + "','" + part + "', '" + Convert.ToDecimal(amt) + "','" + currency + "','', '" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate )";
                        if (QueryExecute(conn_if, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                }
                else
                {
                    //엑셀데이타 값이 없는것은 제외:
                    if ( (yyyymm == "" || yyyymm == null || yyyymm == "&nbsp;") || (item_group == "" || item_group == null || item_group == "&nbsp;")
                        || (part == "" || part == null || part == "&nbsp;") || (amt == "" || amt == null || amt == "&nbsp;")
                        || (currency == "" || currency == null || currency == "&nbsp;")) //값체크
                        MessageBox.ShowMessage( i.ToString() + "번째ROW 칼럼 값이 이상합니다. ", this);                    
                    else
                    {
                        sql = " insert into t_device_amt values( '" + yyyymm + "' , '" + item_group + "','" + part + "', '" + Convert.ToDecimal(amt) + "','" + currency + "','', '" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate )";
                        if (QueryExecute(conn_if, sql, "") > 0)
                            chk_save_yn += 1;
                    }
                }

            }
            if (GridView1.Rows.Count == chk_save_yn)
            {
                MessageBox.ShowMessage("저장되었습니다", this);
            }
            else
            {
                MessageBox.ShowMessage("저장에 문제가 있습니다. 데이타를 확인해 보시기 바랍니다.", this);
            }
            GridView1.DataSource = null;
            GridView1.DataBind();

        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            //그리드 뷰를 초기화 한다.
            GridView1.DataSource = null;
            GridView1.DataBind();
            GridView1.Visible = false;
        }        

        protected void ddlSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            //엑셀 sheet별 데이타 보여주기
            ViewExcelSheets(HiddenField_filePath.Value, HiddenField_extension.Value, "YES");
        }

       

        protected void btn_device_amt_view_Click(object sender, EventArgs e)
        {
            //초기화
            FpSpread_amt.DataSource = null;
            FpSpread_amt.DataBind();
            //updatepanel_amt.Update();

            string frdt, todt;
            if (str_fr_dt.Text == null || str_fr_dt.Text == "")
            {
                MessageBox.ShowMessage("조회시작일을 입력해주세요.", this);
            }
            else if (str_to_dt.Text == null || str_to_dt.Text == "")
            {
                MessageBox.ShowMessage("조회종료일을 입력해주세요.", this);
            }
            else
            {
                frdt = str_fr_dt.Text.Substring(0, 6);
                todt = str_to_dt.Text.Substring(0, 6);
                string sql = "select '' status ,yyyymm, item_group, part, amt,currency from t_device_amt where yyyymm between '" + frdt + "' and '" + todt + "' ";

                sqlAdapter1 = new OracleDataAdapter(sql, conn_if);

                sqlAdapter1.Fill(ds, "ds");

                FpSpread_amt.DataSource = ds;
                FpSpread_amt.DataBind();
                //FpSpread_amt.ActiveSheetView.Columns[4].Visible = false;

            }           
            
        }

        protected void FpSpread_amt_UpdateCommand(object sender, FarPoint.Web.Spread.SpreadCommandEventArgs e)
        {
            int colcnt;
            int i;
            int r = (int)e.CommandArgument;
            colcnt = e.EditValues.Count - 1;
            

            for (i = 0; i <= colcnt; i++)
            {
                if (!object.ReferenceEquals(e.EditValues[i], FarPoint.Web.Spread.FpSpread.Unchanged))
                {
                    string sql = "" ;
                    string chang_data = e.EditValues[i].ToString();
                    Session["column_index"] = i.ToString();
                    //세션 초기화
                    if (i == 0) //상태
                        Session["column_first"] = e.EditValues[i].ToString();

                    if (i == 1) //년월
                    {
                        Session["yyyymm"] = e.EditValues[i].ToString();
                        sql = " update t_device_amt set yyyymm = '"+e.EditValues[i].ToString()+"'  ";
                    }
                    if (i == 2) //품목그룹
                    {
                        Session["item_gp"] = e.EditValues[i].ToString();
                        sql = " update t_device_amt set item_group = '" + e.EditValues[i].ToString() + "'  ";
                    }

                    if (i == 3) //품목
                    {
                        Session["part"] = e.EditValues[i].ToString();
                        sql = " update t_device_amt set part = '" + e.EditValues[i].ToString() + "'  ";
                    }

                    if (i == 4) //금액
                    {
                        Session["amt"] = e.EditValues[i].ToString();
                        sql = " update t_device_amt set amt = " + Convert.ToDecimal( e.EditValues[i].ToString()) + "  ";
                    }

                    if (i == 5) //환율
                    {
                        Session["currency"] = e.EditValues[i].ToString();
                        sql = " update t_device_amt set currency = '" + e.EditValues[i].ToString() + "'  ";
                    }

                    //수정시 최초 값을 산출한다. 
                    if (Session["column_first"].ToString() == "수정" && i != 0)
                    {
                        Session["old_yyyymm"] = FpSpread_amt.Sheets[0].Cells[r, 1].Text;
                        Session["old_itemgp"] = FpSpread_amt.Sheets[0].Cells[r, 2].Text;
                        Session["old_part"] = FpSpread_amt.Sheets[0].Cells[r, 3].Text;
                        Session["old_amt"] = FpSpread_amt.Sheets[0].Cells[r, 4].Text;
                        Session["old_currency"] = FpSpread_amt.Sheets[0].Cells[r, 4].Text;

                        if (Session["old_yyyymm"].ToString() != "" && Session["old_itemgp"].ToString() != "" && Session["old_part"].ToString() != "" && Session["old_amt"].ToString() != "" && Session["old_currency"].ToString() != "")
                        {
                            sql = sql + " where yyyymm = '" + Session["old_yyyymm"].ToString() + "'  and item_group = '" + Session["old_itemgp"].ToString() + "' and part = '" + Session["old_part"].ToString() + "' ";

                            QueryExecute(conn_if, sql, "");
                        }
                    }                                  
                }
            }
        }

        protected void btn_spread_save_Click(object sender, EventArgs e)
        {
            FpSpread_amt.SaveChanges();

            //updatepanel_amt.Update();
            for (int i = 0; i < FpSpread_amt.Rows.Count; i++)
            {
                if (FpSpread_amt.Sheets[0].Cells[i, 0].Text == "수정")
                {

                    string new_yyyymm = "", new_item_gp = "", new_part = "", new_amt = "", new_currency = "";

                    new_yyyymm = FpSpread_amt.Sheets[0].Cells[i, 1].Text;
                    new_item_gp = FpSpread_amt.Sheets[0].Cells[i, 2].Text;
                    new_part = FpSpread_amt.Sheets[0].Cells[i, 3].Text;
                    new_amt = FpSpread_amt.Sheets[0].Cells[i, 4].Text;
                    new_currency = FpSpread_amt.Sheets[0].Cells[i, 5].Text;

                    if (new_yyyymm == null || new_yyyymm == "")
                        MessageBox.ShowMessage("기준년월을 입력해주세요.", this);
                    else if (new_item_gp == null || new_item_gp == "")
                        MessageBox.ShowMessage("품목그룹을 입력해주세요.", this);
                    else if (new_part == null || new_part == "")
                        MessageBox.ShowMessage("디바이스를 입력해주세요.", this);
                    else if (new_amt == null || new_amt == "")
                        MessageBox.ShowMessage("단가를 입력해주세요.", this);
                    else if (new_currency == null || new_currency == "")
                        MessageBox.ShowMessage("환율을 입력해주세요.", this);
                    else
                    {
                        //old값이 없으면 저장
                        if (Session["old_yyyymm"].ToString() == "" && Session["old_itemgp"].ToString() == "" && Session["old_part"].ToString() == "" && Session["old_amt"].ToString() == "" && Session["old_currency"].ToString() == "")
                        {
                            string sql = "insert into t_device_amt values ('" + new_yyyymm + "','" + new_item_gp + "','" + new_part + "','" + Convert.ToDecimal(new_amt) + "','" + new_currency + "' ,'', '" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate  ) ";

                            if (QueryExecute(conn_if, sql, "") > 0)
                            {
                                MessageBox.ShowMessage("저장되었습니다.", this);
                                FpSpread_amt.Sheets[0].SetValue(i, 0, ""); //저장완료
                            }
                            else
                                MessageBox.ShowMessage("저장 오류. 관리자에게 문의하여 주세요.", this);
                        }
                        //else
                        //{
                        //    // old 값이 있으면 업데이트
                        //    string sql = "update t_device_amt set yyyymm = '" + new_yyyymm + "', item_group = '" + new_item_gp + "', part = '" + new_part + "', amt = '" + Convert.ToDecimal(new_amt) + "' , updt_user_id = '" + Session["User"].ToString() + "', updt_dt = sysdate   " +
                        //                 " where yyyymm = '" + Session["old_yyyymm"].ToString() + "'  and item_group = '" + Session["old_itemgp"].ToString() + "' and part = '" + Session["old_part"].ToString() + "' ";

                        //    if (QueryExecute(conn_if, sql, "") > 0)
                        //    {
                        //        FpSpread_amt.Sheets[0].SetValue(i, 0, ""); //저장완료
                        //        MessageBox.ShowMessage("수정되었습니다.", this);
                        //    }
                        //    else
                        //        MessageBox.ShowMessage("수정 오류. 관리자에게 문의하여 주세요.", this);
                        //}
                    }

                    //chk.EditValues[i];
                }
                
                if (FpSpread_amt.Sheets[0].Cells[i, 0].Text == "입력")
                {
                    //string new_index = "", new_yyyymm = "", new_item_gp = "", new_part = "", new_amt = "";

                    ////변경된 인덱스를 가져온다. 
                    //FpSpread_amt.SaveChanges();
                    
                    //new_yyyymm = Session["yyyymm"].ToString();
                    //new_item_gp = Session["item_gp"].ToString();
                    //new_part = Session["part"].ToString();
                    //new_amt = Session["amt"].ToString();
                    //if (new_yyyymm == null || new_yyyymm == "")
                    //    MessageBox.ShowMessage("에러: 기준년월이 없습니다.", this);
                    //else if (new_item_gp == null || new_item_gp == "")
                    //    MessageBox.ShowMessage("에러: 품목그룹이 없습니다.", this);
                    //else if (new_part == null || new_part == "")
                    //    MessageBox.ShowMessage("에러: 디바이스명이 없습니다.", this);
                    //else if (new_amt == null || new_amt == "")
                    //    MessageBox.ShowMessage("에러: 단가가 없습니다.", this);
                    //else
                    //{
                    //    string sql = "insert into t_device_amt values ('" + new_yyyymm + "','" + new_item_gp + "','" + new_part + "','" + Convert.ToDecimal(new_amt) + "', '', '" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate  ) ";

                    //    if (QueryExecute(conn_if, sql, "") > 0)
                    //    {
                    //        MessageBox.ShowMessage("저장되었습니다.", this);
                    //        FpSpread_amt.Sheets[0].SetValue(i, 0, ""); //저장완료
                    //    }
                    //    else
                    //        MessageBox.ShowMessage("저장 오류. 관리자에게 문의하여 주세요.", this);
                    //}

                }
                if (FpSpread_amt.Sheets[0].Cells[i, 0].Text == "삭제")
                {
                    string yyyymm = "", item_gp = "", part = "";
                    yyyymm = FpSpread_amt.Sheets[0].Cells[i, 1].Text;
                    item_gp = FpSpread_amt.Sheets[0].Cells[i, 2].Text;
                    part = FpSpread_amt.Sheets[0].Cells[i, 3].Text;

                    string sql = "delete t_device_amt where yyyymm = '" + yyyymm + "'  and item_group = '" + item_gp + "' and part = '" + part + "' ";

                    if (QueryExecute(conn_if, sql, "") > 0)
                    {
                        MessageBox.ShowMessage("삭제되었습니다.", this);
                        FpSpread_amt.Sheets[0].SetValue(i, 0, ""); //저장완료
                    }
                    else
                        MessageBox.ShowMessage("삭제 오류. 관리자에게 문의하여 주세요.", this);
                }

            }
            btn_device_amt_view_Click(null, null);
        }

        protected void btn_amt_insert_Click(object sender, EventArgs e)
        {
            //1행을 추가한다. 
            FpSpread_amt.Sheets[0].AddRows(FpSpread_amt.Sheets[0].RowCount, 1);
            FpSpread_amt.Sheets[0].SetValue(FpSpread_amt.Sheets[0].RowCount -1, 0, "입력");
            
        }

        protected void btn_amt_delete_Click(object sender, EventArgs e)
        {
            //선택된 여러 row를 확인한다.
            System.Collections.IEnumerator enu = FpSpread_amt.ActiveSheetView.SelectionModel.GetEnumerator();
            FarPoint.Web.Spread.Model.CellRange cr;

            while (enu.MoveNext())
            {
                cr = ((FarPoint.Web.Spread.Model.CellRange)(enu.Current));
                int a = FpSpread_amt.Sheets[0].ActiveRow;
                //FpSpread2.Sheets[0].Rows.Remove(cr.Row, cr.RowCount);
                for (int i = 0; i < cr.RowCount; i++)
                {
                    FpSpread_amt.Sheets[0].SetValue(cr.Row + i, 0, "삭제");                    
                }
            }

            //FpSpread_amt.Sheets[0].SetValue(FpSpread_amt.Sheets[0].ActiveRow, 0, "삭제");
        }

        private void fpsread_data_dbexe(string type, string sql, int row)
        {
            if (type == "update")
            {
                string new_yyyymm = "", new_item_gp = "", new_part = "", new_amt = "", new_currency = ""; ;

                new_yyyymm = FpSpread_amt.Sheets[0].Cells[row, 1].Text;
                new_item_gp = FpSpread_amt.Sheets[0].Cells[row, 2].Text;
                new_part = FpSpread_amt.Sheets[0].Cells[row, 3].Text;
                new_amt = FpSpread_amt.Sheets[0].Cells[row, 4].Text;

                if (new_yyyymm == null || new_yyyymm == "")
                    MessageBox.ShowMessage("기준년월을 입력해주세요.", this);
                else if (new_item_gp == null || new_item_gp == "")
                    MessageBox.ShowMessage("품목그룹을 입력해주세요.", this);
                else if (new_part == null || new_part == "")
                    MessageBox.ShowMessage("디바이스를 입력해주세요.", this);
                else if (new_amt == null || new_amt == "")
                    MessageBox.ShowMessage("단가를 입력해주세요.", this);
                else if (new_currency == null || new_currency == "")
                    MessageBox.ShowMessage("환율을 입력해주세요.", this);
                else
                {
                    //old값이 없으면 저장
                    if (Session["old_yyyymm"].ToString() == "" && Session["old_itemgp"].ToString() == "" && Session["old_part"].ToString() == "" && Session["old_amt"].ToString() == "" && Session["old_currency"].ToString() == "")
                    {
                        sql = "insert into t_device_amt values ('" + new_yyyymm + "','" + new_item_gp + "','" + new_part + "','" + Convert.ToDecimal(new_amt) + "','" + new_currency + "' , '', '" + Session["User"].ToString() + "', sysdate, '" + Session["User"].ToString() + "', sysdate  ) ";

                        if (QueryExecute(conn_if, sql, "") > 0)
                        {
                            MessageBox.ShowMessage("저장되었습니다.", this);
                            FpSpread_amt.Sheets[0].SetValue(row, 0, ""); //저장완료
                        }
                        else
                            MessageBox.ShowMessage("저장 오류. 관리자에게 문의하여 주세요.", this);
                    }
                    else
                    {
                       
                        // old 값이 있으면 업데이트
                        sql = "update t_device_amt set yyyymm = '" + new_yyyymm + "', item_group = '" + new_item_gp + "', part = '" + new_part + "', amt = '" + Convert.ToDecimal(new_amt) + "' currency = '" + new_currency + "', updt_user_id = '" + Session["User"].ToString() + "', updt_dt = sysdate   " +
                                    " where yyyymm = '" + Session["old_yyyymm"].ToString() + "'  and item_group = '" + Session["old_itemgp"].ToString() + "' and part = '" + Session["old_part"].ToString() + "' ";

                        if (QueryExecute(conn_if, sql, "") > 0)
                        {
                            FpSpread_amt.Sheets[0].SetValue(row, 0, ""); //저장완료
                            MessageBox.ShowMessage("수정되었습니다.", this);
                        }
                        else
                            MessageBox.ShowMessage("수정 오류. 관리자에게 문의하여 주세요.", this);
                    }
                }
            }
            if (type == "insert")
            {

            }
            if (type == "delete")
            {
            }
        }

       
    }
}