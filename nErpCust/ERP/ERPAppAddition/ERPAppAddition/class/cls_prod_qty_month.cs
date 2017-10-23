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
using ERPAppAddition.ERPAddition.CM.CM_C2001;

public class cls_prod_qty_month
{
    string strConn = ConfigurationManager.AppSettings["connectionKey"];

    string sql_cust_cd;

    OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_CCUBE_MIGHTY"].ConnectionString);

    OracleCommand cmd = new OracleCommand();
    OracleDataReader dr;

    public DataTable fetch(string fr_dt, string to_dt)
    {
        string sql
    = "SELECT /*+ NO_CPU_COSTING */ YYYYMM,  " +
        "       'QTY' M_GUBUN,  " +
        "       '1-PROD' M_GUBUN_DETAIL,  " +
        "       '생산' M_GUBUN_DETAIL_NM,  " +
        "       SUM(DDI_QTY) DDI_VAL,  " +
        "       SUM(WLP_QTY) WLP_VAL,  " +
        "       SUM(PTEST_QTY) P_TEST_VAL,  " +
        "       SUM(COG_QTY) COG_VAL,  " +
        "       SUM(COF_QTY) COF_VAL,  " +
        "       SUM(FT_QTY) F_TEST_VAL,  " +
        "       SUM(WLCSP_QTY) WLCSP_VAL,  " +
        "       SUM(BUMP_QTY) twelve_val,  " +
        "       SUM(SANG_QTY) sangpum_val, sysdate isrt_dt ,'10' seq  " +
        "  FROM (SELECT SUBSTR('" + fr_dt + "', 1, 6) YYYYMM,  " +
        "               DECODE(DEVICE, 'DDI', MONTH_QTY, 0) DDI_QTY,  " +
        "               DECODE(DEVICE, 'WLP', MONTH_QTY, 0) WLP_QTY,  " +
        "               DECODE(DEVICE, 'P-TEST', MONTH_QTY, 0) PTEST_QTY,  " +
        "               DECODE(DEVICE, 'COG', MONTH_QTY, 0) COG_QTY,  " +
        "               DECODE(DEVICE, 'TAB', MONTH_QTY, 0) COF_QTY,  " +
        "               DECODE(DEVICE, 'TAB', MONTH_QTY, 0) FT_QTY,  " +
        "               DECODE(DEVICE, 'WLCSP', MONTH_QTY, 0) WLCSP_QTY,  " +
        "               DECODE(DEVICE, '12BUMP', MONTH_QTY, 0) BUMP_QTY,  " +
        "               0 SANG_QTY  " +
        "          FROM (SELECT DEVICE,  " +
        "                       SUM(SUM_PRODUCT_MONTH) AS MONTH_QTY  " +
        "                  FROM (SELECT DEVICE,  " +
        "                               SUM(SUM_PRODUCT_MONTH) AS SUM_PRODUCT_MONTH  " +
        "                          FROM (SELECT PLANT,  " +
        "                                       DEVICE,  " +
        "                                       CUSTOMER,  " +
        "                                       SUM_PRODUCT_MONTH  " +
        "                                  FROM (SELECT PLANT,  " +
        "                                               'DDI' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               SUM(WH1.OPER_OUT_QTY1) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH1  " +
        "                                         WHERE WH1.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH1.OPERATION <> WH1.TO_OPERATION  " +
        "                                           AND WH1.TRANSACTION = 'SPOU'  " +
        "                                           AND WH1.OPERATION = '4900'  " +
        "                                           AND WH1.REWORK = 'N'  " +
        "                                           AND (WH1.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH1.ACCOUNT_CODE IN (SELECT SC1.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC1  " +
        "                                                                          WHERE SC1.PLANT = WH1.PLANT  " +
        "                                                                            AND SC1.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC1.SYSCODE_GROUP = 'CHARGE'  " +
        "                                                                          GROUP BY SC1.SYSCODE_NAME))  " +
        "                                           AND WH1.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH1.QTY_UNIT_1 = 'SLS'  " +
        "                                           AND WH1.CREATE_CODE IN ('B', 'BA', 'BALM', 'BM')  " +
        "                                           AND WH1.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               '12BUMP' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               SUM(WH2.OPER_OUT_QTY1) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH2  " +
        "                                         WHERE WH2.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH2.OPERATION <> WH2.TO_OPERATION  " +
        "                                           AND WH2.OPERATION = 'F040'  " +
        "                                           AND WH2.REWORK = 'N'  " +
        "                                           AND (WH2.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH2.ACCOUNT_CODE IN (SELECT SC2.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC2  " +
        "                                                                          WHERE SC2.PLANT = WH2.PLANT  " +
        "                                                                            AND SC2.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC2.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH2.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH2.QTY_UNIT_1 = 'SLS'  " +
        "                                           AND WH2.CREATE_CODE IN ('WLCSP', 'WG', 'WGO', 'WS', 'WSO', 'WSLD')  " +
        "                                           AND WH2.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                           AND WH2.ROUTESET <> 'WLTS-OS-PTE'  " +
        "                                           AND WH2.PART <> '12TESTPART'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'DDI' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               SUM(WH3.OPER_OUT_QTY1) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH3  " +
        "                                         WHERE WH3.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH3.OPERATION <> WH3.TO_OPERATION  " +
        "                                           AND WH3.TRANSACTION = 'SPOU'  " +
        "                                           AND WH3.OPERATION = '4000'  " +
        "                                           AND WH3.REWORK = 'N'  " +
        "                                           AND (WH3.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH3.ACCOUNT_CODE IN (SELECT SC3.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC3  " +
        "                                                                          WHERE SC3.PLANT = WH3.PLANT  " +
        "                                                                            AND SC3.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC3.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH3.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH3.QTY_UNIT_1 = 'SLS'  " +
        "                                           AND WH3.CREATE_CODE IN ('BP', 'BPLM')  " +
        "                                           AND WH3.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'WLP' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               SUM(WH4.OPER_OUT_QTY1) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH4  " +
        "                                         WHERE WH4.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH4.OPERATION <> WH4.TO_OPERATION  " +
        "                                           AND WH4.OPERATION = '4900'  " +
        "                                           AND WH4.REWORK = 'N'  " +
        "                                           AND (WH4.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH4.ACCOUNT_CODE IN (SELECT SC4.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC4  " +
        "                                                                          WHERE SC4.PLANT = WH4.PLANT  " +
        "                                                                            AND SC4.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC4.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH4.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH4.QTY_UNIT_1 = 'SLS'  " +
        "                                           AND WH4.CREATE_CODE IN ('WLCSP', 'WG', 'WGO', 'WS', 'WSO', 'WSLD')  " +
        "                                           AND WH4.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'P-TEST' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               SUM(WH5.OPER_OUT_QTY1) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH5  " +
        "                                         WHERE WH5.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH5.OPERATION <> WH5.TO_OPERATION  " +
        "                                           AND WH5.REWORK = 'N'  " +
        "                                           AND WH5.OPERATION = '4900'  " +
        "                                           AND (WH5.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH5.ACCOUNT_CODE IN (SELECT SC5.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC5  " +
        "                                                                          WHERE SC5.PLANT = WH5.PLANT  " +
        "                                                                            AND SC5.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC5.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH5.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH5.QTY_UNIT_1 = 'SLS'  " +
        "                                           AND WH5.CREATE_CODE IN ('BP', 'BPLM', 'PR', 'PRLM')  " +
        "                                           AND WH5.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'COG' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               ROUND(SUM(WH6.OPER_OUT_QTY1) / 1000, 0) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH6  " +
        "                                         WHERE WH6.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH6.OPERATION <> WH6.TO_OPERATION  " +
        "                                           AND WH6.REWORK = 'N'  " +
        "                                           AND WH6.OPERATION = '8900'  " +
        "                                           AND (WH6.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH6.ACCOUNT_CODE IN (SELECT SC6.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC6  " +
        "                                                                          WHERE SC6.PLANT = WH6.PLANT  " +
        "                                                                            AND SC6.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC6.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH6.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH6.QTY_UNIT_1 = 'PCS'  " +
        "                                           AND WH6.CREATE_CODE = 'COG'  " +
        "                                           AND WH6.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'TAB' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               ROUND(SUM(WH7.OPER_OUT_QTY1) / 1000, 0) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH7  " +
        "                                         WHERE WH7.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH7.OPERATION <> WH7.TO_OPERATION  " +
        "                                           AND WH7.REWORK = 'N'  " +
        "                                           AND WH7.OPERATION = '8900'  " +
        "                                           AND (WH7.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH7.ACCOUNT_CODE IN (SELECT SC7.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC7  " +
        "                                                                          WHERE SC7.PLANT = WH7.PLANT  " +
        "                                                                            AND SC7.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC7.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH7.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH7.QTY_UNIT_1 = 'PCS'  " +
        "                                           AND WH7.CREATE_CODE = 'TCP'  " +
        "                                           AND WH7.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER  " +
        "                                        UNION ALL  " +
        "                                        SELECT PLANT,  " +
        "                                               'WLCSP' AS DEVICE,  " +
        "                                               CUSTOMER,  " +
        "                                               ROUND(SUM(WH8.OPER_OUT_QTY1) / 1000, 0) AS SUM_PRODUCT_MONTH  " +
        "                                          FROM WIPHST WH8  " +
        "                                         WHERE WH8.PLANT = 'CCUBEDIGITAL'  " +
        "                                           AND WH8.OPERATION <> WH8.TO_OPERATION  " +
        "                                           AND WH8.REWORK = 'N'  " +
        "                                           AND WH8.OPERATION = '8900'  " +
        "                                           AND (WH8.ACCOUNT_CODE IS NULL  " +
        "                                                 OR WH8.ACCOUNT_CODE IN (SELECT SC8.SYSCODE_NAME  " +
        "                                                                           FROM SYSCODEDATA SC8  " +
        "                                                                          WHERE SC8.PLANT = WH8.PLANT  " +
        "                                                                            AND SC8.SYSTABLE_NAME = 'SCRAP_REASON'  " +
        "                                                                            AND SC8.SYSCODE_GROUP = 'CHARGE'))  " +
        "                                           AND WH8.LOT_SUB_TYPE = 'NONE'  " +
        "                                           AND WH8.QTY_UNIT_1 = 'PCS'  " +
        "                                           AND WH8.CREATE_CODE IN ('CRCN', 'CFLM', 'WLCSP', 'CG', 'CS')  " +
        "                                           AND WH8.REPORT_DATE BETWEEN '" + fr_dt + "' AND '" + to_dt + "'  " +
        "                                         GROUP BY PLANT, CUSTOMER))  " +
        "                         GROUP BY DEVICE  " +
        "                        UNION ALL  " +
        "                        SELECT CODE_NAME AS DEVICE,  " +
        "                               0 AS SUM_PRODUCT_MONTH  " +
        "                          FROM USERCODEDATA  " +
        "                         WHERE PLANT = 'CCUBEDIGITAL'  " +
        "                           AND TABLE_NAME = 'DEVICE_GTYPE')  " +
        "                 GROUP BY ROLLUP (DEVICE)))  " +
        " GROUP BY YYYYMM ";

        ds_cm_c2001 ds = new ds_cm_c2001();

        conn.Open();
        cmd = conn.CreateCommand();
        cmd.CommandType = CommandType.Text;
        cmd.CommandText = sql;

        try
        {
            dr = cmd.ExecuteReader();
            ds.Tables[0].Load(dr);
            conn.Close();
        }
        catch (Exception ex)
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        return ds.Tables[0];
    }
    
}
