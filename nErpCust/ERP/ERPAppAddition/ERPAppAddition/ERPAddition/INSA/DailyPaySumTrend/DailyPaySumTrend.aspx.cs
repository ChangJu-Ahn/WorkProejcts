using System.Web;
using System.Data;
using System;
using System.Web.UI.WebControls;
using System.Text;
using SRL.UserControls;


namespace ERPAppAddition.ERPAddition.INSA.DailyPaySumTrend
{
    public partial class DailyPaySumTrend : System.Web.UI.Page
    {
        #region "Functions"

        private void SetDefaultValue()
        {
            string Depart = "";

            try
            {
                if (dr_dept.SelectedItem.Text == "Semi")
                {
                    Depart = "Semi";

                    GV.gOraCon.Open();
                    GF.SetPartIDCtrl(GV.gOraCon, mcc_dr_Part, Depart);
                    GF.SetOprGroupIDCtrl(GV.gOraCon, mcc_dr_Oprgrp, Depart);
                    GF.SetAreaIDCtrl(GV.gOraCon, mcc_dr_Area, Depart);
                }
                else if (dr_dept.SelectedItem.Text == "Display")
                {
                    Depart = "Display";

                    GV.gOraCon.Open();
                    mcc_dr_Part.ClearAll();
                    GF.SetOprGroupIDCtrl(GV.gOraCon, mcc_dr_Oprgrp, Depart);
                    GF.SetAreaIDCtrl(GV.gOraCon, mcc_dr_Area, Depart);
                }
                else
                {
                    mcc_dr_Part.ClearAll();
                    mcc_dr_Oprgrp.ClearAll();
                    mcc_dr_Area.ClearAll();
                }
            }
            catch
            {
                GV.gOraCon.Close();
            }
            finally
            {
                GV.gOraCon.Close();
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = "NEPES&NEPES_DISPLAY";
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        private bool GetDefaultValue(ref string strCrntDT, ref string strDeptSemi, ref string strDeptDisp, ref string strAreaID, ref string strOperID, ref string strPartID, ref string strAreaID_Disp, ref string strOperID_Disp,
                                     ref string strAreaID_SAWON, ref string strOperID_SAWON, ref string strPartID_SAWON, ref string strAreaID_Disp1, ref string strOperID_Disp1, ref DateTime To_date, ref DateTime E_date, ref DateTime S_date)
        {
            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi") strDeptSemi = ", 'Semi' AS DEPT \n";
            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi") strDeptDisp = ", 'Display' AS DEPT \n";

            if (dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Area.Count > 0) strAreaID = @"AND B.ASSIGN_GROUP IN ( SELECT DISTINCT B.SUB_CODE AS AREA_CODE
                if (mcc_dr_Area.Count > 0 && mcc_dr_Area.Text != "") strAreaID = @"AND B.ASSIGN_GROUP IN ( SELECT DISTINCT B.SUB_CODE AS AREA_CODE
                                                                                      FROM SAWON_WORK_TIME_INF@CCUBE A, PM000STD@CCUBE B
                                                                                    WHERE 1=1
                                                                                    AND B.SUB_CODE = A.SAWON_GROUP AND B.MAIN_CODE = '00067' AND B.SUB_NAME NOT LIKE 'HRMS%' AND B.SUB_CODE <> '00000'
                                                                                    AND B.SUB_NAME IN (" + mcc_dr_Area.SQLText.Trim() + ")) \n";

                //if (mcc_dr_Oprgrp.Count > 0) strOperID = @"AND B.ASSIGN_BAY IN ( SELECT DISTINCT B.SUB_CODE AS OPER_CODE
                if (mcc_dr_Oprgrp.Count > 0 && mcc_dr_Oprgrp.Text != "") strOperID = @"AND B.ASSIGN_BAY IN ( SELECT DISTINCT B.SUB_CODE AS OPER_CODE
                                                                                       FROM SAWON_WORK_TIME_INF@CCUBE A,  PM000STD@CCUBE B 
                                                                                     WHERE 1=1 
                                                                                       AND B.SUB_CODE = A.SAWON_BAY AND B.MAIN_CODE = '00066' AND B.SUB_NAME NOT LIKE 'HRMS%' AND B.SUB_CODE <> '00000'
                                                                                       AND B.SUB_NAME IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ")) \n";

                //if (mcc_dr_Part.Count > 0) strPartID = @"AND B.ASSIGN_PART IN ( SELECT DISTINCT MANAGE1 AS PART FROM PM000STD@CCUBE
                if (mcc_dr_Part.Count > 0 && mcc_dr_Part.Text != "") strPartID = @"AND B.ASSIGN_PART IN ( SELECT DISTINCT MANAGE1 AS PART FROM PM000STD@CCUBE
                                                                                    WHERE 1=1 AND MAIN_CODE = '00066' AND MANAGE2 IS NOT NULL AND SUB_CODE <> '000000' AND MANAGE1 <> 'Part 구분'
                                                                                    AND MANAGE1 IN (" + mcc_dr_Part.SQLText.Trim() + ")) \n";

                //if (mcc_dr_Area.Count > 0) strAreaID_SAWON = @"AND SAWON_GROUP IN ( SELECT DISTINCT B.SUB_CODE AS AREA_CODE
                if (mcc_dr_Area.Count > 0 && mcc_dr_Area.Text != "") strAreaID_SAWON = @"AND SAWON_GROUP IN ( SELECT DISTINCT B.SUB_CODE AS AREA_CODE
                                                                                      FROM SAWON_WORK_TIME_INF@CCUBE A, PM000STD@CCUBE B
                                                                                    WHERE 1=1
                                                                                    AND B.SUB_CODE = A.SAWON_GROUP AND B.MAIN_CODE = '00067' AND B.SUB_NAME NOT LIKE 'HRMS%' AND B.SUB_CODE <> '00000'
                                                                                    AND B.SUB_NAME IN (" + mcc_dr_Area.SQLText.Trim() + ")) \n";

                //if (mcc_dr_Oprgrp.Count > 0) strOperID_SAWON = @"AND SAWON_BAY IN ( SELECT DISTINCT B.SUB_CODE AS OPER_CODE
                if (mcc_dr_Oprgrp.Count > 0 && mcc_dr_Oprgrp.Text != "") strOperID_SAWON = @"AND SAWON_BAY IN ( SELECT DISTINCT B.SUB_CODE AS OPER_CODE
                                                                                       FROM SAWON_WORK_TIME_INF@CCUBE A,  PM000STD@CCUBE B 
                                                                                     WHERE 1=1 
                                                                                       AND B.SUB_CODE = A.SAWON_BAY AND B.MAIN_CODE = '00066' AND B.SUB_NAME NOT LIKE 'HRMS%' AND B.SUB_CODE <> '00000'
                                                                                       AND B.SUB_NAME IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ")) \n";

                //if (mcc_dr_Part.Count > 0) strPartID_SAWON = @"AND SAWON_PART IN ( SELECT DISTINCT MANAGE1 AS PART FROM PM000STD@CCUBE
                if (mcc_dr_Part.Count > 0 && mcc_dr_Part.Text != "") strPartID_SAWON = @"AND SAWON_PART IN ( SELECT DISTINCT MANAGE1 AS PART FROM PM000STD@CCUBE
                                                                                    WHERE 1=1 AND MAIN_CODE = '00066' AND MANAGE2 IS NOT NULL AND SUB_CODE <> '000000' AND MANAGE1 <> 'Part 구분'
                                                                                    AND MANAGE1 IN (" + mcc_dr_Part.SQLText.Trim() + ")) \n";
            }

            if (dr_dept.SelectedValue == "Display")
            {
                //if (mcc_dr_Area.Count > 0) strAreaID_Disp = "AND B.AREA IN (" + mcc_dr_Area.SQLText.Trim() + ") \n";
                //if (mcc_dr_Oprgrp.Count > 0) strOperID_Disp = "AND B.OPER_GROUP IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ") \n";
                //if (mcc_dr_Area.Count > 0) strAreaID_Disp1 = "AND AREA IN (" + mcc_dr_Area.SQLText.Trim() + ") \n";
                //if (mcc_dr_Oprgrp.Count > 0) strOperID_Disp1 = "AND OPER_GROUP IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ") \n";
                if (mcc_dr_Area.Count > 0 && mcc_dr_Area.Text != "") strAreaID_Disp = "AND B.AREA IN (" + mcc_dr_Area.SQLText.Trim() + ") \n";
                if (mcc_dr_Oprgrp.Count > 0 && mcc_dr_Oprgrp.Text != "") strOperID_Disp = "AND B.OPER_GROUP IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ") \n";
                if (mcc_dr_Area.Count > 0 && mcc_dr_Area.Text != "") strAreaID_Disp1 = "AND AREA IN (" + mcc_dr_Area.SQLText.Trim() + ") \n";
                if (mcc_dr_Oprgrp.Count > 0 && mcc_dr_Oprgrp.Text != "") strOperID_Disp1 = "AND OPER_GROUP IN (" + mcc_dr_Oprgrp.SQLText.Trim() + ") \n";
            }

            if (txtCrntDT.Text.Trim().Length > 0)
            {
                strCrntDT = txtCrntDT.Text;
                To_date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM"));
                S_date = Convert.ToDateTime(strCrntDT + "-01");
                E_date = new DateTime();

                if (strCrntDT == To_date.ToString("yyyy-MM"))
                {
                    E_date = (Convert.ToDateTime(DateTime.Now.ToString("yy-MM-dd")));
                }
                else
                {
                    char[] point = { '-' };
                    string[] stDay = strCrntDT.Split(point);
                    int year = Convert.ToInt32(stDay[0]);
                    int monthly = Convert.ToInt32(stDay[1]);

                    E_date = Convert.ToDateTime(strCrntDT + "-" + Convert.ToString((DateTime.DaysInMonth(year, monthly))));
                }
            }
            return true;
        }

        private string MakeQuery(string strCrntDT, string strDeptSemi, string strDeptDisp, string strAreaID, string strOperID, string strPartID, string strAreaID_SAWON, string strOperID_SAWON, string strPartID_SAWON,
                                 string strAreaID_Disp, string strOperID_Disp, string strAreaID_Disp1, string strOperID_Disp1, string strSQL_Prod, DateTime To_date, DateTime E_date, DateTime S_date)
        {
            string strSQL;
            StringBuilder sb_semi = new StringBuilder();
            StringBuilder strProd_SQL = new StringBuilder();
            StringBuilder strProd_SQL_Bump = new StringBuilder();
            StringBuilder strProd_SQL_PTest = new StringBuilder();

            // 여기서부터 반도체 쿼리
            sb_semi.AppendLine("SELECT");
            sb_semi.AppendLine("    TO_CHAR(TO_DATE(WORK_DATE),'YY/MM/DD') AS WORK_DATE ");
            sb_semi.AppendLine("   ,TO_CHAR(SUM(CTT))CTT");
            sb_semi.AppendLine("   ,TO_CHAR(SUM(CT))CT");
            sb_semi.AppendLine("   ,TO_CHAR ('') as aa");
            sb_semi.AppendLine("   ,SUM(TIME_SUM)TIME_SUM");
            sb_semi.AppendLine("   ,SUM(NWRK_TIME)NWRK_TIME");
            sb_semi.AppendLine("   ,SUM(OW_TIME)OW_TIME");
            sb_semi.AppendLine("   ,SUM(NIGHT_TIME)NIGHT_TIME");
            sb_semi.AppendLine("   ,SUM(PAY_SUM) / 1000 AS PAY_SUM");
            sb_semi.AppendLine("   ,SUM(BASE_PAY) / 1000 AS BASE_PAY");
            sb_semi.AppendLine("   ,SUM(OW_TIME_PAY) / 1000 AS OW_TIME_PAY");
            sb_semi.AppendLine("   ,SUM(HOLWRK_OW_TIME_PAY) / 1000 AS HOLWRK_OW_TIME_PAY");
            sb_semi.AppendLine("   ,SUM(HOLWRK_TIME_PAY) / 1000 AS HOLWRK_TIME_PAY");
            sb_semi.AppendLine("   ,SUM(NIGHT_TIME_PAY) / 1000 AS NIGHT_TIME_PAY");
            sb_semi.AppendLine("   ,SUM(ALLOWANCE_2_PAY) / 1000 AS ALLOWANCE_2_PAY");
            sb_semi.AppendLine("   ,SUM(RESIGN_PAY) / 1000 AS RESIGN_PAY");
            sb_semi.AppendLine("   ,SUM(INSURANCE_PAY) / 1000 AS INSURANCE_PAY");
            sb_semi.AppendLine("   ,SUM(MAINT_PAY) / 1000 AS MAINT_PAY");
            sb_semi.AppendLine("   ,SUM(ETC_PAY) / 1000 AS ETC_PAY");
            sb_semi.AppendLine("   " + strDeptSemi + " FROM ");
            sb_semi.AppendLine("   (");
            sb_semi.AppendLine("   SELECT A.WORK_DATE");
            sb_semi.AppendLine("   ,A.CTT,CT,V_CT");
            sb_semi.AppendLine("   ,TIME_SUM,NWRK_TIME,OW_TIME,NIGHT_TIME,PAY_SUM,BASE_PAY,OW_TIME_PAY,HOLWRK_OW_TIME_PAY,HOLWRK_TIME_PAY,NIGHT_TIME_PAY,ALLOWANCE_2_PAY,RESIGN_PAY,INSURANCE_PAY,MAINT_PAY,ETC_PAY");
            sb_semi.AppendLine("   FROM");
            sb_semi.AppendLine("   (");
            //sb_semi.AppendLine("   SELECT WORK_DATE,CTT");
            //sb_semi.AppendLine("   FROM");
            //sb_semi.AppendLine("     (");
            sb_semi.AppendLine(" WITH SAWON_TOTAL_INWON AS ");
            sb_semi.AppendLine("     (");

            // 3일미만 입사자 제외요청 관련 이전쿼리 주석 후 신규 쿼리 적용 (2015.11.12, ahncj, 생산팀 윤미화 대리님 요청)
            //for (DateTime day = S_date; day <= E_date; day = day.AddDays(1))
            //{
            //    sb_semi.AppendLine("     (");
            //    sb_semi.AppendLine("     SELECT '" + day.ToString("yyyy-MM-dd") + "' AS WORK_DATE,COUNT(A.SAWON_NO)AS CTT");
            //    sb_semi.AppendLine("     FROM (SELECT * FROM SAWON_INF@CCUBE WHERE SAWON_JOIN_DATE <= '" + day.ToString("yyyy-MM-dd") + "' AND ( SAWON_QUIT_DATE IS NULL OR SAWON_QUIT_DATE >= '" + day.ToString("yyyy-MM-dd") + "' ))A");
            //    sb_semi.AppendLine("     ,(");
            //    sb_semi.AppendLine("     SELECT A.*");
            //    sb_semi.AppendLine("     FROM (SELECT * FROM SAWON_ASSIGN_INF@CCUBE WHERE DEL_FLAG IS NULL)A");
            //    sb_semi.AppendLine("     ,(SELECT SAWON_NO,MAX(ASSIGN_SEQ) ASSIGN_SEQ FROM SAWON_ASSIGN_INF@CCUBE WHERE DEL_FLAG IS NULL AND ASSIGN_DATE <= '" + day.ToString("yyyy-MM-dd") + "' GROUP BY SAWON_NO)B");
            //    sb_semi.AppendLine("     WHERE A.SAWON_NO = B.SAWON_NO");
            //    sb_semi.AppendLine("     AND A.ASSIGN_SEQ = B.ASSIGN_SEQ");
            //    sb_semi.AppendLine("     )B");
            //    sb_semi.AppendLine("     WHERE A.SAWON_NO = B.SAWON_NO");
            //    sb_semi.AppendLine(" " + strAreaID + strOperID + strPartID + " )");

            //    if (day < E_date)
            //    {
            //        sb_semi.AppendLine("UNION ALL");
            //    }
            //}

            for (DateTime day = S_date; day <= E_date; day = day.AddDays(1))
            {
                sb_semi.AppendLine("     SELECT '" + day.ToString("yyyy-MM-dd") + "' AS WORK_DATE,COUNT(A.SAWON_NO)AS CTT");
                sb_semi.AppendLine("     FROM ( ");
                sb_semi.AppendLine("     SELECT * FROM SAWON_INF@CCUBE  ");
                sb_semi.AppendLine("     WHERE SAWON_JOIN_DATE <= '" + day.ToString("yyyy-MM-dd") + "' ");
                sb_semi.AppendLine("     AND ( SAWON_QUIT_DATE IS NULL OR SAWON_QUIT_DATE >= '" + day.ToString("yyyy-MM-dd") + "') ");
                sb_semi.AppendLine("     AND SAWON_DIVISION='반도체' ");
                sb_semi.AppendLine("     AND (SAWON_QUIT_GBN <> '입사취소' OR SAWON_QUIT_GBN IS NULL) )A ");
                sb_semi.AppendLine("     ,(");
                sb_semi.AppendLine("     SELECT A.*");
                sb_semi.AppendLine("     FROM (SELECT * FROM SAWON_ASSIGN_INF@CCUBE WHERE DEL_FLAG IS NULL)A");
                sb_semi.AppendLine("     ,(SELECT SAWON_NO,MAX(ASSIGN_SEQ) ASSIGN_SEQ FROM SAWON_ASSIGN_INF@CCUBE WHERE DEL_FLAG IS NULL AND ASSIGN_DATE <= '" + day.ToString("yyyy-MM-dd") + "' GROUP BY SAWON_NO)B");
                sb_semi.AppendLine("     WHERE A.SAWON_NO = B.SAWON_NO");
                sb_semi.AppendLine("     AND A.ASSIGN_SEQ = B.ASSIGN_SEQ");
                sb_semi.AppendLine("     )B");
                sb_semi.AppendLine("     WHERE A.SAWON_NO = B.SAWON_NO(+)");
                sb_semi.AppendLine(" " + strAreaID + strOperID + strPartID + " ");

                if (day < E_date)
                {
                    sb_semi.AppendLine("UNION ALL");
                }
            }

            sb_semi.AppendLine(" ) ");
            sb_semi.AppendLine(" SELECT * FROM SAWON_TOTAL_INWON ");
            sb_semi.AppendLine(" 		   )A,");
            sb_semi.AppendLine(" (");
            sb_semi.AppendLine(" SELECT WORK_DATE");
            sb_semi.AppendLine(" ,SUM(W_CNT) AS CT");
            sb_semi.AppendLine(" ,SUM(V_CNT) AS V_CT");
            sb_semi.AppendLine(" ,ROUND(SUM(NW_TIME) + SUM(OW_TIME) + SUM(HOLWRK_TIME) + SUM(HOLWRK_OW_TIME),1) AS TIME_SUM");
            sb_semi.AppendLine(" ,(SUM(NW_TIME) + SUM(HOLWRK_TIME)) AS NWRK_TIME");
            sb_semi.AppendLine(" ,(SUM(OW_TIME) + SUM(HOLWRK_OW_TIME)) AS OW_TIME");
            sb_semi.AppendLine(" ,SUM(NIGHT_TIME) AS NIGHT_TIME");
            sb_semi.AppendLine(" ,ROUND( SUM(ROUND(BASE_PAY, 0))");
            sb_semi.AppendLine(" + SUM(OW_TIME_PAY)");
            sb_semi.AppendLine(" + SUM(HOLWRK_OW_TIME_PAY)");
            sb_semi.AppendLine(" + SUM(HOLWRK_TIME_PAY)");
            sb_semi.AppendLine(" + SUM(NIGHT_TIME_PAY)");
            sb_semi.AppendLine(" + SUM(ALLOWANCE_2_PAY)");
            sb_semi.AppendLine(" + ((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) / 12)");
            sb_semi.AppendLine(" + ((SUM(BASE_PAY) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY))  * 0.09979)");
            sb_semi.AppendLine(" + SUM(MAINT_PAY)");
            sb_semi.AppendLine(" + SUM(DATA_1_PANGONG_PAY),0) AS PAY_SUM");
            sb_semi.AppendLine(" ,SUM(ROUND(BASE_PAY, 0)) AS BASE_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(OW_TIME_PAY), 0) AS OW_TIME_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(HOLWRK_OW_TIME_PAY), 0) AS HOLWRK_OW_TIME_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(HOLWRK_TIME_PAY) , 0)AS HOLWRK_TIME_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(NIGHT_TIME_PAY), 0) AS NIGHT_TIME_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(ALLOWANCE_2_PAY), 0) AS ALLOWANCE_2_PAY");
            sb_semi.AppendLine(" ,ROUND((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) / 12, 0) AS RESIGN_PAY");
            sb_semi.AppendLine(" ,ROUND( ((SUM(BASE_PAY) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY))  * 0.09979), 0) AS INSURANCE_PAY");
            sb_semi.AppendLine(" ,SUM(MAINT_PAY) AS MAINT_PAY");
            sb_semi.AppendLine(" ,ROUND(SUM(DATA_1_PANGONG_PAY),0) AS ETC_PAY");
            sb_semi.AppendLine(" FROM");
            sb_semi.AppendLine(" (");
            sb_semi.AppendLine(" SELECT C.SAWON_NO");
            sb_semi.AppendLine(" ,WORK_DATE");
            sb_semi.AppendLine(" ,CASE WHEN NW_TIME IS NOT NULL OR NW_TIME > 0 THEN 1 ELSE 0 END W_CNT");
            sb_semi.AppendLine(" ,CASE WHEN (NW_TIME IS NULL AND VACATION_KIND IS NOT NULL) OR VACATION_KIND ='결근' THEN 1 ELSE 0 END V_CNT");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,NULL,NW_TIME),0)NW_TIME");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,'O',NW_TIME),0)HOLWRK_TIME");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,NULL,OW_TIME),0)OW_TIME");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,'O',OW_TIME),0)HOLWRK_OW_TIME");
            sb_semi.AppendLine(" ,NVL(NIGHT_TIME,0)NIGHT_TIME");
            sb_semi.AppendLine(" ,VACATION_KIND");
            sb_semi.AppendLine(" ,NVL(LATE_TIME,0)LATE_TIME");
            sb_semi.AppendLine(" ,NVL(EARLY_TIME,0)EARLY_TIME");
            sb_semi.AppendLine(" ,D.SAWON_PAY_GBN");
            sb_semi.AppendLine(" ,TIME_PAY");
            //sb_semi.AppendLine(" ,NVL(CASE WHEN SAWON_PAY_GBN ='시급' THEN TIME_PAY * 209 / ( TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD') )");
            //sb_semi.AppendLine(" ELSE (MONTH_PAY + EXTRA_PAY) / ( TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD') )");
            //sb_semi.AppendLine(" END,0) AS BASE_PAY");
            sb_semi.AppendLine(" ,CASE WHEN VACATION_KIND IN ('출산휴가', '병가', '무급', '생휴') THEN 0 --실제 무급이 발생되는 카테고리는 기본급을 0원 처리 (좌측 4개 내용 외 기본급 발생) ");
            sb_semi.AppendLine("                                                               ELSE ( NVL(CASE WHEN SAWON_PAY_GBN ='시급' THEN TIME_PAY * 209 / ( TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD') )");
            sb_semi.AppendLine("                                                                                                          ELSE (MONTH_PAY + EXTRA_PAY) / ( TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD') )");
            sb_semi.AppendLine("                                                                          END,0) ) ");
            sb_semi.AppendLine("  END AS BASE_PAY   -- 기본급 (시급 * 209 * 해당 월 MAX일수 ) ");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,NULL, NVL(OW_TIME,0) * TIME_PAY * 1.5),0) OW_TIME_PAY");
            sb_semi.AppendLine(" ,NVL(CASE WHEN SAWON_PAY_GBN ='시급' AND HOLIDAY_WORK ='O' THEN NVL(NW_TIME,0) * TIME_PAY * 1.5");
            sb_semi.AppendLine(" WHEN SAWON_PAY_GBN ='월급' AND HOLIDAY_WORK ='O' THEN (NVL(NW_TIME,0) / 4 ) * 30000");
            sb_semi.AppendLine(" END,0) AS HOLWRK_TIME_PAY");
            sb_semi.AppendLine(" ,NVL(DECODE(HOLIDAY_WORK,'O', NVL(OW_TIME,0) * TIME_PAY * 2),0) HOLWRK_OW_TIME_PAY");
            sb_semi.AppendLine(" ,NVL((NVL(NIGHT_TIME,0) * TIME_PAY * 0.5),0) NIGHT_TIME_PAY");
            sb_semi.AppendLine(" ,NVL((NVL(LATE_TIME,0) + NVL(EARLY_TIME,0)) * TIME_PAY,0) LATE_EARLYLEAVE_PAY");
            sb_semi.AppendLine(" ,NVL(DECODE(SAWON_PAY_GBN,'시급', (TIME_PAY * 209 )/12 / ( TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD') )* (CASE WHEN NW_TIME IS NOT NULL OR NW_TIME > 0 THEN 1 ELSE 0 END)),0) ALLOWANCE_2_PAY");
            sb_semi.AppendLine(" ,ROUND( (115000 / (TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD'))),0) AS MAINT_PAY");
            sb_semi.AppendLine(" ,NVL((NVL(LEVEL_PAY,0) +NVL( ETC_PAY,0)) / (TO_CHAR(LAST_DAY('" + S_date.ToString("yyyy-MM-dd") + "'), 'DD')),0) DATA_1_PANGONG_PAY");
            //sb_semi.AppendLine(" FROM (SELECT * FROM SAWON_WORK_TIME_INF@CCUBE WHERE WORK_DATE BETWEEN TO_CHAR('" + S_date.ToString("yyyy-MM-dd") + "') AND TO_CHAR('" + E_date.ToString("yyyy-MM-dd") + "')");
            sb_semi.AppendLine(" FROM ( SELECT A.* FROM SAWON_WORK_TIME_INF@CCUBE A INNER JOIN SAWON_INF@CCUBE B ON A.SAWON_NO = B.SAWON_NO AND (B.SAWON_QUIT_DATE <= TO_CHAR('" + E_date.ToString("yyyy-MM-dd") + "') OR B.SAWON_QUIT_DATE IS NULL) AND (B.SAWON_QUIT_GBN <> '입사취소' OR B.SAWON_QUIT_GBN IS NULL) AND SAWON_DIVISION = '반도체' ");
            sb_semi.AppendLine(" WHERE A.WORK_DATE BETWEEN TO_CHAR('" + S_date.ToString("yyyy-MM-dd") + "') AND TO_CHAR('" + E_date.ToString("yyyy-MM-dd") + "') ");

            sb_semi.AppendLine(" " + strAreaID_SAWON + strOperID_SAWON + strPartID_SAWON + ") C");

            sb_semi.AppendLine(" ,(SELECT A.*");
            sb_semi.AppendLine(" FROM");
            sb_semi.AppendLine(" SAWON_PAY_INF@ccube A,");
            sb_semi.AppendLine(" (");
            sb_semi.AppendLine(" SELECT SAWON_NO,MAX(SEQ)SEQ");
            sb_semi.AppendLine(" FROM SAWON_PAY_INF@CCUBE");
            sb_semi.AppendLine(" WHERE DEL_FLAG IS NULL");
            //sb_semi.AppendLine(" AND (E_DATE IS NULL OR E_DATE > '" + S_date.ToString("yyyy-MM-dd") + "') AND S_DATE <= '" + S_date.ToString("yyyy-MM-dd") + "'"); 중도입사자가 생길 경우 시급을 찾아올 수 없는 문제점 관련 수정
            sb_semi.AppendLine(" AND (E_DATE IS NULL OR E_DATE > '" + S_date.ToString("yyyy-MM-dd") + "') AND S_DATE <= '" + E_date.ToString("yyyy-MM-dd") + "'");
            sb_semi.AppendLine(" GROUP BY SAWON_NO");
            sb_semi.AppendLine(" )B ");
            sb_semi.AppendLine(" WHERE A.SAWON_NO = B.SAWON_NO");
            sb_semi.AppendLine(" AND A.SEQ = B.SEQ)D");
            sb_semi.AppendLine(" WHERE C.SAWON_NO = D.SAWON_NO(+)");
            sb_semi.AppendLine(" )");
            sb_semi.AppendLine(" GROUP BY WORK_DATE");
            sb_semi.AppendLine(" )B");
            sb_semi.AppendLine(" WHERE A.WORK_DATE =B.WORK_DATE(+)");
            sb_semi.AppendLine(" )");
            sb_semi.AppendLine(" GROUP BY WORK_DATE");
            // 여기까지 반도체 쿼리

            // 여기서부터 디스플레이 쿼리
            strSQL = "   SELECT                                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "     VAL_1.*                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "   , VAL_2.TARGET_PAY / 1000 AS TARGET_PAY                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "   , VAL_2.HOLWRK_FLAG AS HOLWRK_FLAG                                                                                                                                                                                                                            \n";

            if (strSQL_Prod.Length > 0)
            {
                strSQL = strSQL + "   , RESLT_1.PROD_QTY AS PROD_QTY                                                                                                                                                                                                                            \n";
            }

            strSQL = strSQL + "   FROM                                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "  (SELECT DISTINCT                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "     WORK_DATE                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "   , SUM(CTT) AS CTT                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "   , SUM(CT) AS CT                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "   , ROUND((SUM(CT) / SUM(CTT)) * 100, 0) AS PROPORTION                                                                                                                                                                                                          \n";
            strSQL = strSQL + "   , SUM(TIME_SUM) AS TIME_SUM                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "   , SUM(NWRK_TIME) AS NWRK_TIME                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "   , SUM(OW_TIME) AS OW_TIME                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "   , SUM(NIGHT_TIME) AS NIGHT_TIME                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "   , SUM(PAY_SUM) AS PAY_SUM                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "   , SUM(BASE_PAY) AS BASE_PAY                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "   , SUM(OW_TIME_PAY) AS OW_TIME_PAY                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "   , SUM(HOLWRK_OW_TIME_PAY) AS HOLWRK_OW_TIME_PAY                                                                                                                                                                                                               \n";
            strSQL = strSQL + "   , SUM(HOLWRK_TIME_PAY) AS HOLWRK_TIME_PAY                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "   , SUM(NIGHT_TIME_PAY) AS NIGHT_TIME_PAY                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "   , SUM(ALLOWANCE_2_PAY) AS ALLOWANCE_2_PAY                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "   , SUM(RESIGN_PAY) AS RESIGN_PAY                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "   , SUM(INSURANCE_PAY) AS INSURANCE_PAY                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "   , SUM(MAINT_PAY) AS MAINT_PAY                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "   , SUM(ETC_PAY) AS ETC_PAY                                                                                                                                                                                                                                     \n";

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                strSQL = strSQL + "   , DEPT                                                                                                                                                                                                                                                    \n";
            }

            strSQL = strSQL + "FROM (                                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                WITH DEFUALT AS (                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "            SELECT                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "        F.WORK_DATE                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "       ,ROUND( SUM(F.NWRK_TIME - F.LATE_EARLYLEAVE) + SUM(F.OW_TIME) + SUM(F.HOLWRK_TIME) + SUM(F.HOLWRK_OW_TIME)  ,1) AS TIME_SUM -- 근태정보합계                                                                                                               \n";
            strSQL = strSQL + "       ,SUM(F.NWRK_TIME - F.LATE_EARLYLEAVE) + SUM(F.HOLWRK_TIME) AS NWRK_TIME -- 정상근무 (정상시간 - 지각조퇴 시간)                                                                                                                                            \n";
            strSQL = strSQL + "       ,SUM(F.OW_TIME) + SUM(F.HOLWRK_OW_TIME) AS OW_TIME --연장                                                                                                                                                                                                 \n";
            strSQL = strSQL + "       ,SUM(F.NIGHT_TIME) AS NIGHT_TIME -- 야간                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "       ,ROUND( SUM(F.BASE_PAY)                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "          + SUM(F.OW_TIME_PAY)                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "          + SUM(F.HOLWRK_OW_TIME_PAY)                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "          + SUM(F.HOLWRK_TIME_PAY)                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "                     + SUM(F.NIGHT_TIME_PAY)                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "                     + SUM(F.ALLOWANCE_2_PAY)                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "                     + ( (SUM(F.BASE_PAY) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) /12 )                                                                                      \n";
            strSQL = strSQL + "                     + (( SUM(F.BASE_PAY) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) * 0.09979)                                                                                 \n";
            strSQL = strSQL + "                     + SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD')) )                                                                                                                                                                          \n";
            strSQL = strSQL + "                     + SUM(F.DATA_1_PANGONG_PAY + F.DATA_1_LEVEL_PAY) ,0) AS PAY_SUM -- 급여정보 합계                                                                                                                                                            \n";
            strSQL = strSQL + "                  ,SUM(ROUND(F.BASE_PAY, 0)) AS BASE_PAY -- 기본급                                                                                                                                                                                               \n";
            strSQL = strSQL + "                  ,SUM(F.OW_TIME_PAY) AS OW_TIME_PAY -- 평일연장                                                                                                                                                                                                 \n";
            strSQL = strSQL + "                  ,SUM(F.HOLWRK_OW_TIME_PAY) AS HOLWRK_OW_TIME_PAY -- 휴일연장                                                                                                                                                                                   \n";
            strSQL = strSQL + "                  ,SUM(F.HOLWRK_TIME_PAY) AS HOLWRK_TIME_PAY -- 휴일근무                                                                                                                                                                                         \n";
            strSQL = strSQL + "                  ,SUM(F.NIGHT_TIME_PAY) AS NIGHT_TIME_PAY -- 야간                                                                                                                                                                                               \n";
            strSQL = strSQL + "                  ,SUM(F.ALLOWANCE_2_PAY) AS ALLOWANCE_2_PAY --상여금                                                                                                                                                                                            \n";
            strSQL = strSQL + "                  ,ROUND(SUM(F.BASE_PAY) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY),0) / 12 AS RESIGN_PAY                                                                        \n";
            strSQL = strSQL + "                  ,ROUND( ((SUM(ROUND(F.BASE_PAY, 0)) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) * 0.09979), 0) AS INSURANCE_PAY                                                \n";
            strSQL = strSQL + "                  ,ROUND(SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD'))),0) AS MAINT_PAY -- 관리비                                                                                                                                               \n";
            strSQL = strSQL + "                  ,ROUND(SUM(F.DATA_1_PANGONG_PAY + F.DATA_1_LEVEL_PAY),0) AS ETC_PAY -- 기타급여                                                                                                                                                                \n";
            strSQL = strSQL + "          FROM (                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "                                     SELECT DISTINCT TIME_DATA.*     -- 하위쿼리의 기본 인적사항 및 시간                                                                                                                                                         \n";
            strSQL = strSQL + "                                              , C.HOURLY_WAGE AS HOURLY_WAGE_PAY -- 개인시급                                                                                                                                                                     \n";
            strSQL = strSQL + "                                              , CASE WHEN TIME_DATA.WORK_GROUP  IN ('제조1파트','제조2파트') THEN C.HOURLY_WAGE * 209 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD') ) * WRK_DAY                                                            \n";
            strSQL = strSQL + "                                                         ELSE (C.MONTHLY_WAGE + C.ALLOWANCE_7) / ( TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD') )                                                                                                             \n";
            strSQL = strSQL + "                                                 END AS BASE_PAY   -- 기본급 (시급 * 209 * 해당 월 MAX일수 * 실제 근무일수)                                                                                                                                      \n";
            strSQL = strSQL + "                                              , (TIME_DATA.OW_TIME * C.HOURLY_WAGE * 1.5) AS OW_TIME_PAY   -- 평일연장수당 ( 시급 * 1.5 * 평일연장근무시간 )                                                                                                     \n";
            strSQL = strSQL + "                                              , (TIME_DATA.HOLWRK_OW_TIME * C.HOURLY_WAGE * 2) AS HOLWRK_OW_TIME_PAY   -- 주말연장수당 ( 시급 * 2 * 주말연장근무시간)                                                                                            \n";
            strSQL = strSQL + "                                              , CASE WHEN  TIME_DATA.WORK_GROUP IN ('제조1파트', '제조2파트') THEN TIME_DATA.HOLWRK_TIME * C.HOURLY_WAGE * 1.5                                                                                                   \n";
            strSQL = strSQL + "                                                         ELSE (TIME_DATA.HOLWRK_TIME / 4 ) * 30000                                                                                                                                                               \n";
            strSQL = strSQL + "                                                END AS HOLWRK_TIME_PAY    -- 특근수당 ( 시급 * 1.5 * 주말근무시간 )                                                                                                                                              \n";
            strSQL = strSQL + "                                              , (TIME_DATA.NIGHT_TIME * C.HOURLY_WAGE * 0.5) AS NIGHT_TIME_PAY   -- 야간수당 ( 야간 근무시간 * 시급 * 0.5 )                                                                                                      \n";
            strSQL = strSQL + "                                              , TIME_DATA.LATE_EARLYLEAVE * C.HOURLY_WAGE AS LATE_EARLYLEAVE_PAY            -- 지각조퇴 차감 ( 시급 * 지각조퇴시간 )                                                                                             \n";
            strSQL = strSQL + "                                              , CASE WHEN TIME_DATA.WORK_GROUP  IN ('제조1파트','제조2파트') THEN (C.HOURLY_WAGE * 209) / 12 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD') ) * WRK_DAY                                                     \n";
            strSQL = strSQL + "                                                          ELSE 0                                                                                                                                                                                                 \n";
            strSQL = strSQL + "                                                 END AS ALLOWANCE_2_PAY          -- 상여금 (개인별 입력수당 )                                                                                                                                                    \n";
            strSQL = strSQL + "                                              ,  TIME_DATA.DATA_1_PANGONG /  (TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD'))  AS DATA_1_PANGONG_PAY       -- 판공비 (기준정보 고정수당 )                                                                       \n";
            strSQL = strSQL + "                                              ,  TIME_DATA.DATA_1_LEVEL /  (TO_CHAR(LAST_DAY('" + strCrntDT + "-01'), 'DD')) AS DATA_1_LEVEL_PAY   --인증수당 ( 기준정보 고정수당 )                                                                              \n";
            strSQL = strSQL + "                                      FROM (                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "                                              SELECT DISTINCT                                                                                                                                                                                                    \n";
            strSQL = strSQL + "                                                       B.WORK_DATE                                                                                                                                                                                               \n";
            strSQL = strSQL + "                                                     , B.VENDOR                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                                                     , B.OPER_GROUP                                                                                                                                                                                              \n";
            strSQL = strSQL + "                                                     , B.AREA                                                                                                                                                                                                    \n";
            strSQL = strSQL + "                                                     , NVL(SUM(CASE WHEN B.ABS_TYPE IN ('휴무', '공가', '연차') OR B.HOLWRK_FLAG = 'Y' THEN 0                                                                                                                    \n";
            strSQL = strSQL + "                                                                    ELSE B.NW_TIME                                                                                                                                                                               \n";
            strSQL = strSQL + "                                                               END),0) AS NWRK_TIME   -- 평일(시간))                                                                                                                                                             \n";
            strSQL = strSQL + "                                                     , CASE B.WORK_GROUP when '1' THEN '제조1파트'                                                                                                                                                               \n";
            strSQL = strSQL + "                                                                                          when '2' THEN '제조2파트'                                                                                                                                              \n";
            strSQL = strSQL + "                                                                                           when 'N' THEN '주간'                                                                                                                                                  \n";
            strSQL = strSQL + "                                                          END AS WORK_GROUP   -- 근무조                                                                                                                                                                          \n";
            strSQL = strSQL + "                                                       , A. EMP_ID -- 사번                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                                       , CASE WHEN (B.NW_TIME = 0 OR B.ABS_TYPE  IN ('무급', '병가')) THEN 0                                                                                                                                     \n";
            strSQL = strSQL + "                                                              ELSE 1                                                                                                                                                                                             \n";
            strSQL = strSQL + "                                                         END AS WRK_DAY   -- 근무일수                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                       , NVL(SUM(B.LATE_TIME + B.EALV_TIME), 0) AS LATE_EARLYLEAVE    -- 지각조퇴(시간)                                                                                                                          \n";
            strSQL = strSQL + "                                                      , CASE WHEN  B.WORK_GROUP IN ('1', '2') THEN NVL(SUM(DECODE(B.HOLWRK_FLAG,'', B.OW_TIME, NULL, B.OW_TIME)), 0)                                                                                             \n";
            strSQL = strSQL + "                                                                  ELSE 0                                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                           END AS OW_TIME   -- 평일연장(시간)                               --                                                                                                                                   \n";
            strSQL = strSQL + "                                                        , CASE WHEN  B.WORK_GROUP IN ('1', '2') THEN  NVL(SUM(DECODE(B.HOLWRK_FLAG,'Y', B.OW_TIME,0)), 0)                                                                                                        \n";
            strSQL = strSQL + "                                                                  ELSE 0                                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                         END AS HOLWRK_OW_TIME   -- 휴일연장(시간)                                                                                                                                                               \n";
            strSQL = strSQL + "                                                        , SUM(CASE WHEN B.HOLWRK_FLAG = 'Y' AND (B.ABS_TYPE IS NULL OR B.ABS_TYPE != '휴무') THEN B.NW_TIME - B.LATE_TIME                                                                                         \n";
            strSQL = strSQL + "                                                                   ELSE 0                                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                              END) AS HOLWRK_TIME   -- 특근(시간)                                                                                                                                                                 \n";
            strSQL = strSQL + "                                                       , CASE WHEN  B.WORK_GROUP IN ('1', '2') THEN NVL(SUM(B.NIGHT_TIME), 0)                                                                                                                                    \n";
            strSQL = strSQL + "                                                                  ELSE 0                                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                           END AS NIGHT_TIME   -- 야간(시간)                                                                                                                                                                     \n";
            strSQL = strSQL + "                                                        , DECODE(D.DATA_1, '', 0,  NULL, 0, D.DATA_1) AS DATA_1_PANGONG   -- 판공비(직급수당)                                                                                                                    \n";
            strSQL = strSQL + "                                                        , DECODE(E.DATA_1, '', 0,  NULL, 0, E.DATA_1) AS DATA_1_LEVEL   --인센티브(인증수당)                                                                                                                     \n";
            strSQL = strSQL + "                                                FROM NHRMWRKDEF A INNER JOIN NHRMWRKIOT B ON A.EMP_ID = B.EMP_ID AND B.NW_TIME >= 1                                                                                                                              \n";
            strSQL = strSQL + "                                                                                  LEFT OUTER JOIN MGCMTBLDAT D ON A.POSITION = D.KEY_2 AND D.FACTORY = 'DISPLAY' AND D.TABLE_NAME = 'ALLOWANCE'                                                                  \n";
            strSQL = strSQL + "                                                                                  LEFT OUTER JOIN MGCMTBLDAT E ON A.CERT_LEVEL = E.KEY_2 AND E.FACTORY = 'DISPLAY' AND E.TABLE_NAME = 'ALLOWANCE'                                                                \n";
            strSQL = strSQL + "                                                  WHERE B.WORK_DATE BETWEEN '" + S_date.ToString("yyyy-MM-dd") + "' AND '" + E_date.ToString("yyyy-MM-dd") + "'     /*이쪽에 각종 조건문을 넣자*/                                                                \n";
            strSQL = strSQL + "                                                        AND A.RESIGN_TYPE <> '입사취소'    /* 퇴사사유 중 재직 3일 미만 (입사취소자) 데이터는 출력하지 않도록 수정 [2015.11.11, ahncj, 생산팀 윤미화 대리 요청] */                                                           \n";

            strSQL = strSQL + strAreaID_Disp + strOperID_Disp;

            strSQL = strSQL + "                                                GROUP BY B.WORK_GROUP, D.DATA_1, E.DATA_1, A.EMP_ID, B.WORK_DATE, B.VENDOR, B.OPER_GROUP, B.AREA, B.NW_TIME, B.ABS_TYPE                                                                                          \n";
            strSQL = strSQL + "                                                )TIME_DATA                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                   INNER JOIN (                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                                                SELECT B.*                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                                 FROM (                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                                                       SELECT MAX(WORK_DATE) AS WORK_DATE, EMP_ID                                                                                                                                                                \n";
            strSQL = strSQL + "                                                        FROM NHRMWRKPAY                                                                                                                                                                                          \n";
            strSQL = strSQL + "                                                         WHERE WORK_DATE <= '" + E_date.ToString("yyyy-MM-dd") + "'                                                                                                                                               \n";
            strSQL = strSQL + "                                                        GROUP BY EMP_ID) A, NHRMWRKPAY B                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                WHERE 1=1                                                                                                                                                                                                        \n";
            strSQL = strSQL + "                                                 AND A.WORK_DATE = B.WORK_DATE                                                                                                                                                                                   \n";
            strSQL = strSQL + "                                                 AND A.EMP_ID = B.EMP_ID) C ON TIME_DATA.EMP_ID = C.EMP_ID                                                                                                                                                       \n";
            strSQL = strSQL + "            ) F                                                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "            GROUP BY                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "             F.WORK_DATE                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "            ORDER BY F.WORK_DATE                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "  )                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  ,                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  HUMAN_TOTAL AS ( -- 재적인원                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "    SELECT                                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "          WORK_DATE                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "        , COUNT(*) AS CTT                                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "    FROM NHRMWRKIOT                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "      WHERE WORK_DATE BETWEEN '" + S_date.ToString("yyyy-MM-dd") + "' AND '" + E_date.ToString("yyyy-MM-dd") + "'                                                                                                                                                                                               \n";
            strSQL = strSQL + strAreaID_Disp1 + strOperID_Disp1;
            strSQL = strSQL + "    GROUP BY WORK_DATE                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "  )                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  ,                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  HUMAN_SUM AS ( -- 출근안한 인원                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "    SELECT                                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "          WORK_DATE                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "        , COUNT(*) AS CT                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "    FROM NHRMWRKIOT                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "      WHERE 1=1                                                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "       AND WORK_DATE BETWEEN '" + S_date.ToString("yyyy-MM-dd") + "' AND '" + E_date.ToString("yyyy-MM-dd") + "'                                                                                                                                                                                                  \n";
            strSQL = strSQL + "      AND ABS_TYPE NOT IN  ('정상', '지각', '반차', '비상', '조퇴', '반공')                                                                                                                                                                                      \n";
            strSQL = strSQL + strAreaID_Disp1 + strOperID_Disp1;
            strSQL = strSQL + "    GROUP BY WORK_DATE                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "  )                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  ,                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "  HUMAN AS ( -- 재적인원 - 근태인원 (재적인원 - 출근안한 인원)                                                                                                                                                                                                   \n";
            strSQL = strSQL + "   SELECT                                                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "       HUMAN_TOTAL.WORK_DATE AS WORK_DATE                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "     , HUMAN_TOTAL.CTT AS CTT                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "     , (HUMAN_TOTAL.CTT - HUMAN_SUM.CT) AS CT                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "     , ROUND(((HUMAN_TOTAL.CTT - HUMAN_SUM.CT) / HUMAN_TOTAL.CTT) * 100, 0) || '%' AS RATIO                                                                                                                                                                      \n";
            strSQL = strSQL + "   FROM HUMAN_TOTAL, HUMAN_SUM                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "     WHERE HUMAN_TOTAL.WORK_DATE = HUMAN_SUM.WORK_DATE(+)                                                                                                                                                                                                        \n";
            strSQL = strSQL + "           )                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + " SELECT                                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "       TO_CHAR(Y.WORK_DATE) AS WORK_DATE                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "     , TO_CHAR(Y.CTT) AS CTT                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "     , TO_CHAR(Y.CT) AS CT                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "     , TO_CHAR(Y.RATIO) AS PROPORTION                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.TIME_SUM) AS TIME_SUM                                                                                                                                                                                                                \n";
            strSQL = strSQL + "      ,TO_NUMBER(Z.NWRK_TIME) AS NWRK_TIME                                                                                                                                                                                                              \n";
            strSQL = strSQL + "      ,TO_NUMBER(Z.OW_TIME) AS OW_TIME                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "      ,TO_NUMBER(Z.NIGHT_TIME) AS NIGHT_TIME                                                                                                                                                                                                            \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.PAY_SUM),0) / 1000 AS PAY_SUM                                                                                                                                                                                                           \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.BASE_PAY),0) / 1000 AS BASE_PAY                                                                                                                                                                                                         \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.OW_TIME_PAY),0) / 1000 AS OW_TIME_PAY                                                                                                                                                                                                   \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.HOLWRK_OW_TIME_PAY),0) / 1000 AS HOLWRK_OW_TIME_PAY                                                                                                                                                                                     \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.HOLWRK_TIME_PAY),0) / 1000 AS HOLWRK_TIME_PAY                                                                                                                                                                                           \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.NIGHT_TIME_PAY),0) / 1000 AS NIGHT_TIME_PAY                                                                                                                                                                                             \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.ALLOWANCE_2_PAY),0) / 1000 AS ALLOWANCE_2_PAY                                                                                                                                                                                           \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.RESIGN_PAY),0) / 1000 AS RESIGN_PAY                                                                                                                                                                                                     \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.INSURANCE_PAY),0) / 1000 AS INSURANCE_PAY                                                                                                                                                                                               \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.MAINT_PAY),0) / 1000 AS MAINT_PAY                                                                                                                                                                                                       \n";
            strSQL = strSQL + "      , ROUND(TO_NUMBER(Z.ETC_PAY),0) / 1000 AS ETC_PAY                                                                                                                                                                                                           \n";
            strSQL = strSQL + strDeptDisp;
            strSQL = strSQL + " FROM DEFUALT Z, HUMAN Y                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + " WHERE 1=1                                                                                                                                                                                                                                                        \n";
            strSQL = strSQL + " AND Z.WORK_DATE = Y.WORK_DATE                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "UNION ALL                                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + sb_semi;

            //if (dr_dept.SelectedValue == "Semi")
            //{
            //    if (strSQL_Prod.Length > 0)
            //    {
            //        strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, NHRMDPYSUM VAL_2, (" + strSQL_Prod + ") RESLT_1                                                                                                                                                                         \n";
            //    }
            //    else if (mcc_dr_Area.SQLText.Length == 0 && mcc_dr_Part.SQLText.Length == 0 && mcc_dr_Oprgrp.SQLText.Length == 0 && strSQL_Prod.Length > 0)
            //    {
            //        strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, NHRMDPYSUM VAL_2                                                                                                                                                                                                       \n";
            //    }
            //   // else

            //}

            if (dr_dept.SelectedValue == "Semi")
            {
                if (mcc_dr_Area.SQLText.Length == 0 && mcc_dr_Part.SQLText.Length == 0 && mcc_dr_Oprgrp.SQLText.Length == 0 && strSQL_Prod.Length == 0)
                {
                    strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, NHRMDPYSUM VAL_2 \n";
                }
                else if (strSQL_Prod.Length > 0)
                {
                    strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, (SELECT WORK_DATE, HOLWRK_FLAG, '0' AS TARGET_PAY FROM NHRMDPYSUM) VAL_2, (" + strSQL_Prod + ") RESLT_1                                                                                                                                                                                                        \n";
                }
                else
                {
                    strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, (SELECT WORK_DATE, HOLWRK_FLAG, '0' AS TARGET_PAY FROM NHRMDPYSUM) VAL_2 \n";
                }
            }
            else if (dr_dept.SelectedValue == "Display")
            {
                strSQL = strSQL + ") GROUP BY WORK_DATE, DEPT) VAL_1, (SELECT WORK_DATE, HOLWRK_FLAG, '0' AS TARGET_PAY FROM NHRMDPYSUM) VAL_2                                                                                                                                                                                                            \n";
            }
            else
            {
                strSQL = strSQL + ") GROUP BY WORK_DATE) VAL_1, (SELECT WORK_DATE, HOLWRK_FLAG, '0' AS TARGET_PAY FROM NHRMDPYSUM) VAL_2                                                                                                                                                                                                                 \n";
            }

            strSQL = strSQL + "WHERE VAL_1.WORK_DATE = VAL_2.WORK_DATE(+)                                                                                                                                                                                                                        \n";

            if (strSQL_Prod.Length > 0)
            {
                strSQL = strSQL + "AND VAL_1.WORK_DATE = RESLT_1.REPORT_DATE(+)                                                                                                                                                                                                                  \n";
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                strSQL = strSQL + "AND DEPT = '" + dr_dept.SelectedValue + "' \n";
            }

            return strSQL;
        }

        private string MakeQuery_Prod(string strSQL_Prod, DateTime S_date, DateTime E_date)
        {
            switch (drp_Prod.SelectedValue)
            {
                case "A":
                    strSQL_Prod = "";
                    break;

                case "B": // 범프 실적쿼리
                    if (drp_Prod_sub.SelectedValue == "ALL")
                    {
                        strSQL_Prod = "                                    SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                    \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                     \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                        \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                    AND ( A.ACCOUNT_CODE IS NULL                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                       AND (                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                               (ROUTESET NOT LIKE '%-T' AND A.OPERATION = B.OUT_OPER)                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                               OR                                                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                               (ROUTESET LIKE '%-T' AND A.OPERATION = '4013')                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                           )                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_1 IN ('BUMP')                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                     AND (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) IN ('6' , '8' )                                 \n";
                        strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "  WHERE(                                                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________C%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________H%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________J%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "           NOT((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE LIKE 'PR%' AND CUSTOMER = 'SILICON MITUS'))    \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "  OR(                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________C%'))                             \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________H%'))                             \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________J%'))                             \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE 'PR%' AND CUSTOMER = 'SILICON MITUS'))       \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "  OR (SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'DDI')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "  OR(                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND PTEST_FLAG = 'N' AND OPERATION = '4013' ))                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "     )                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "           GROUP BY REPORT_DATE                                                                                                                                   \n";
                        break;
                    }

                    else if (drp_Prod_sub.SelectedValue == "DDI")
                    {
                        strSQL_Prod = "     SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "         SUM (PROD_QTY) AS PROD_QTY                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "     FROM                                                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "     (                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "           SELECT A.PLANT                                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                , B.SUB_PLANT_1                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                , B.SUB_PLANT_2                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                , A.CUSTOMER                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                , A.PART AS DEVICE                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                           \n";
                        strSQL_Prod = strSQL_Prod + "                , B.PROC_TYPE                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                            \n";
                        strSQL_Prod = strSQL_Prod + "                , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                               \n";
                        strSQL_Prod = strSQL_Prod + "                , A.REPORT_DATE                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                , A.OPERATION                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                , B.PTEST_FLAG                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "            FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "           WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "             AND A.PLANT = B.PLANT                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "             AND A.CREATE_CODE = B.PROC_TYPE                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "             AND A.REWORK ='N'                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "             AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "             AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "             AND ( A.ACCOUNT_CODE IS NULL                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                   OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                                            FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                                           WHERE SCD.PLANT = A.PLANT                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                                             AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                                             AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                                        )                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                   )                                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "              AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "              AND A.CREATE_CODE  <> 'INK'                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                AND (                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                        (ROUTESET NOT LIKE '%-T' AND A.OPERATION = B.OUT_OPER)                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                        OR                                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                        (ROUTESET LIKE '%-T' AND A.OPERATION = '4013')                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                    )                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "              AND B.SUB_PLANT_1 IN ('BUMP')                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "              AND B.SUB_PLANT_2 IN ('DDI')                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "              AND A.PART_TYPE IN ('P','D')                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "              AND (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) IN ('6' , '8' )                                        \n";
                        strSQL_Prod = strSQL_Prod + "     )                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "    WHERE (SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'DDI')                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "    GROUP BY REPORT_DATE                                                                                                                                          \n";
                        break;
                    }

                    else
                    {
                        strSQL_Prod = strSQL_Prod + "                                SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                             \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                                \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                    AND ( A.ACCOUNT_CODE IS NULL                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       AND (                                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                               (ROUTESET NOT LIKE '%-T' AND A.OPERATION = B.OUT_OPER)                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                               OR                                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                               (ROUTESET LIKE '%-T' AND A.OPERATION = '4013')                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                           )                                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_1 IN ('BUMP')                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_2 IN ('WLP')                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                     AND (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) IN ('6' , '8' )                                         \n";
                        strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "  WHERE(                                                                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________C%'))                                 \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________H%'))                                 \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE NOT LIKE '__________J%'))                                 \n";
                        strSQL_Prod = strSQL_Prod + "           AND                                                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "           NOT((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'N' AND OPERATION = '4900' AND ROUTE LIKE 'PR%' AND CUSTOMER = 'SILICON MITUS'))            \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "  OR(                                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________C%'))                                     \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________H%'))                                     \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________J%'))                                     \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4000' AND ROUTE LIKE '__________PR%' AND CUSTOMER = 'SILICON MITUS'))     \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "  OR(                                                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'BUMP' AND SUB_PLANT_2 = 'WLP'  AND PTEST_FLAG = 'N' AND OPERATION = '4013'))                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "  GROUP BY REPORT_DATE                                                                                                                                                    \n";
                        break;
                    }


                case "C": // P-TEST 실적쿼리                                                                                                              
                    if (drp_Prod_sub.SelectedValue == "ALL")
                    {
                        strSQL_Prod = "                                   SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "              SUM (PROD_QTY) AS PROD_QTY                                                                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "          FROM                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "          (                                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                SELECT A.PLANT                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                     , B.SUB_PLANT_1                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                     , B.SUB_PLANT_2                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                     , A.CUSTOMER                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                     , A.PART AS DEVICE                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                     , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                     \n";
                        strSQL_Prod = strSQL_Prod + "                     , B.PROC_TYPE                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                      \n";
                        strSQL_Prod = strSQL_Prod + "                     , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                         \n";
                        strSQL_Prod = strSQL_Prod + "                     , A.REPORT_DATE                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                     , A.OPERATION                                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     , B.PTEST_FLAG                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                 FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                  AND A.PLANT = B.PLANT                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                  AND A.CREATE_CODE = B.PROC_TYPE                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                  AND A.REWORK ='N'                                                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "                  AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                  AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                  AND ( A.ACCOUNT_CODE IS NULL                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                        OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "                                                 FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                                                WHERE SCD.PLANT = A.PLANT                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                                                  AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                         \n";
                        strSQL_Prod = strSQL_Prod + "                                                  AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                                                  AND NOT SCD.SYSCODE_NAME = 'ER-CUST_PLT'                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                                             )                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                        )                                                                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.CREATE_CODE  <> 'INK'                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.OPERATION <> A.TO_OPERATION                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.OPERATION = B.OUT_OPER                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                   AND B.SUB_PLANT_1 IN ('P-TEST')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.PART_TYPE IN ('P','D')                                                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "          )                                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "WHERE(                                                                                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "         ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________C%'))                            \n";
                        strSQL_Prod = strSQL_Prod + "         OR                                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "         ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________H%'))                            \n";
                        strSQL_Prod = strSQL_Prod + "         OR                                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "         ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________J%'))                            \n";
                        strSQL_Prod = strSQL_Prod + "         OR                                                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "         ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE 'PR%' AND CUSTOMER = 'SILICON MITUS'))      \n";
                        strSQL_Prod = strSQL_Prod + ")                                                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "            OR (SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'DDI' )                                                                                                 \n";
                        strSQL_Prod = strSQL_Prod + "         GROUP BY REPORT_DATE                                                                                                                                    \n";
                        break;
                    }

                    else if (drp_Prod_sub.SelectedValue == "DDI")
                    {
                        strSQL_Prod = "                                   SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                  \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                   \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                      \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "                   AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                     AND ( A.ACCOUNT_CODE IS NULL                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND NOT SCD.SYSCODE_NAME = 'ER-CUST_PLT'                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION <> A.TO_OPERATION                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION = B.OUT_OPER                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_1 IN ('P-TEST')                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_2 IN ('DDI')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "           WHERE (SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'DDI' )                                                                                              \n";
                        strSQL_Prod = strSQL_Prod + "  GROUP BY REPORT_DATE                                                                                                                                          \n";
                        break;
                    }

                    else
                    {
                        strSQL_Prod = "                                   SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                  \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                   \n";
                        strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                      \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                           \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                        strSQL_Prod = strSQL_Prod + "                    AND ( A.ACCOUNT_CODE IS NULL                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                        \n";
                        strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                      \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                                                    AND NOT SCD.SYSCODE_NAME = 'ER-CUST_PLT'                                                                    \n";
                        strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                     \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION <> A.TO_OPERATION                                                                                                          \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION = B.OUT_OPER                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_1 IN ('P-TEST')                                                                                                            \n";
                        strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_2 IN ('WLP')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                               \n";
                        strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "  WHERE (                                                                                                                                                       \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________C%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________H%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE '__________J%'))                         \n";
                        strSQL_Prod = strSQL_Prod + "           OR                                                                                                                                                   \n";
                        strSQL_Prod = strSQL_Prod + "           ((SUB_PLANT_1 = 'P-TEST' AND SUB_PLANT_2 = 'WLP' AND PTEST_FLAG = 'Y' AND OPERATION = '4900' AND ROUTE LIKE 'PR%' AND CUSTOMER = 'SILICON MITUS'))   \n";
                        strSQL_Prod = strSQL_Prod + "  )                                                                                                                                                             \n";
                        strSQL_Prod = strSQL_Prod + "           GROUP BY REPORT_DATE                                                                                                                                 \n";
                        break;
                    }


                case "D": // TAB 쿼리                                                                                                                  
                    strSQL_Prod = "                          SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                         \n";
                    strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                      \n";
                    strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                       \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                  \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                   \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                      \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                           \n";
                    strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                       \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                           \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "                    AND ( A.ACCOUNT_CODE IS NULL                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                        \n";
                    strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                     \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION <> A.TO_OPERATION                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION = B.OUT_OPER                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "                     AND B.SUB_PLANT_1 IN ('TAB')                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                   \n";
                    strSQL_Prod = strSQL_Prod + "           GROUP BY REPORT_DATE                                                                                                                                 \n";
                    break;


                case "E": //COG, WLCSP 실적                                                                                                                                                                                                                                                                                                                                                                      
                    strSQL_Prod = "                            SELECT TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                                                       \n";
                    strSQL_Prod = strSQL_Prod + "                SUM (PROD_QTY) AS PROD_QTY                                                                                                                      \n";
                    strSQL_Prod = strSQL_Prod + "            FROM                                                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "            (                                                                                                                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                  SELECT A.PLANT                                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_1                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.SUB_PLANT_2                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.CUSTOMER                                                                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.PART AS DEVICE                                                                                                                       \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT WAFER_DIA FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS WAFER_DIA                                  \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.PROC_TYPE                                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT ROUTESET FROM MIGHTY.PARTROUTESET@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS ROUTE                                   \n";
                    strSQL_Prod = strSQL_Prod + "                       , (SELECT PKGTYPE FROM MIGHTY.PARTSPEC@CCUBE WHERE PLANT = A.PLANT AND PART_ID = A.PART) AS PKGTYPE                                      \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.REPORT_DATE                                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.OPERATION                                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , A.OPER_OUT_QTY1 AS PROD_QTY                                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                       , B.PTEST_FLAG                                                                                                                           \n";
                    strSQL_Prod = strSQL_Prod + "                   FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                  WHERE A.PLANT = 'CCUBEDIGITAL'                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.PLANT = B.PLANT                                                                                                                       \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.CREATE_CODE = B.PROC_TYPE                                                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REWORK ='N'                                                                                                                           \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "                    AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "                    AND ( A.ACCOUNT_CODE IS NULL                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                          OR A.ACCOUNT_CODE IN ( SELECT SCD.SYSCODE_NAME                                                                                        \n";
                    strSQL_Prod = strSQL_Prod + "                                                   FROM MIGHTY.SYSCODEDATA@CCUBE SCD                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                                  WHERE SCD.PLANT = A.PLANT                                                                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSTABLE_NAME = 'SCRAP_REASON'                                                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                                    AND SCD.SYSCODE_GROUP = 'CHARGE'                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                               )                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                          )                                                                                                                                     \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.LOT_SUB_TYPE  = 'NONE'                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.CREATE_CODE  <> 'INK'                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION <> A.TO_OPERATION                                                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "                     AND A.OPERATION = B.OUT_OPER                                                                                                               \n";

                    if (drp_Prod_sub.SelectedValue == "DDI")
                    {
                        strSQL_Prod = strSQL_Prod + " AND B.SUB_PLANT_1 IN ('COG')  \n";
                    }
                    else if (drp_Prod_sub.SelectedValue == "WLP")
                    {
                        strSQL_Prod = strSQL_Prod + " AND B.SUB_PLANT_1 IN ('WLCSP')  \n";
                    }
                    else
                    {
                        strSQL_Prod = strSQL_Prod + " AND B.SUB_PLANT_1 IN ('COG','WLCSP')  \n";
                    }

                    strSQL_Prod = strSQL_Prod + "                     AND A.PART_TYPE IN ('P','D')                                                                                                               \n";
                    strSQL_Prod = strSQL_Prod + "            )                                                                                                                                                   \n";
                    strSQL_Prod = strSQL_Prod + "           GROUP BY REPORT_DATE                                                                                                                                 \n";
                    break;


                case "F": //12인치범프실적
                    strSQL_Prod = "  SELECT                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "      TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE,                                \n";
                    strSQL_Prod = strSQL_Prod + "      SUM(PROD_QTY) AS PROD_QTY                                                                           \n";
                    strSQL_Prod = strSQL_Prod + "  FROM (                                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "              SELECT A.PLANT,                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                                         B.SUB_PLANT_1,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         B.SUB_PLANT_2,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         A.REPORT_DATE,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         A.OPER_OUT_QTY1 AS PROD_QTY                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                    FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B         \n";
                    strSQL_Prod = strSQL_Prod + "                                   WHERE     A.PLANT = 'CCUBEDIGITAL'                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.PLANT = B.PLANT                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.CREATE_CODE = B.PROC_TYPE                                  \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REWORK = 'N'                                               \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND (A.ACCOUNT_CODE IS NULL                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                              OR A.ACCOUNT_CODE IN                                        \n";
                    strSQL_Prod = strSQL_Prod + "                                                    (SELECT SCD.SYSCODE_NAME                              \n";
                    strSQL_Prod = strSQL_Prod + "                                                       FROM MIGHTY.SYSCODEDATA@CCUBE SCD                  \n";
                    strSQL_Prod = strSQL_Prod + "                                                      WHERE SCD.PLANT = A.PLANT                           \n";
                    strSQL_Prod = strSQL_Prod + "                                                            AND SCD.SYSTABLE_NAME =                       \n";
                    strSQL_Prod = strSQL_Prod + "                                                                   'SCRAP_REASON'                         \n";
                    strSQL_Prod = strSQL_Prod + "                                                            AND SCD.SYSCODE_GROUP = 'CHARGE'))            \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.LOT_SUB_TYPE = 'NONE'                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.CREATE_CODE <> 'INK'                                       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.OPERATION <> A.TO_OPERATION                                \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.OPERATION = B.OUT_OPER                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND B.SUB_PLANT_1 = '12BUMP'                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND B.SUB_PLANT_1 IN ('12BUMP')  \n";

                    if (drp_Prod_sub.SelectedValue == "DDI")
                    {
                        strSQL_Prod = strSQL_Prod + " AND B.SUB_PLANT_2 IN ('DDI')  \n";
                    }
                    else if (drp_Prod_sub.SelectedValue == "WLP")
                    {
                        strSQL_Prod = strSQL_Prod + " AND B.SUB_PLANT_2 IN ('WLP')  \n";
                    }

                    strSQL_Prod = strSQL_Prod + "                                         AND A.PART_TYPE IN ('P', 'D')                                    \n";
                    strSQL_Prod = strSQL_Prod + "             )                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "  GROUP BY REPORT_DATE                                                                                    \n";
                    break;

                case "G": //FOWLP 실적쿼리
                    strSQL_Prod = "  SELECT                                                                                                                \n";
                    strSQL_Prod = strSQL_Prod + "       TO_CHAR(TO_DATE(REPORT_DATE, 'YYYYMMDD'),'YY/MM/DD') AS REPORT_DATE                                \n";
                    strSQL_Prod = strSQL_Prod + "     , SUM(PROD_QTY) AS PROD_QTY                                                                          \n";
                    strSQL_Prod = strSQL_Prod + "  FROM (                                                                                                  \n";
                    strSQL_Prod = strSQL_Prod + "              SELECT A.PLANT,                                                                             \n";
                    strSQL_Prod = strSQL_Prod + "                                         B.SUB_PLANT_1,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         B.SUB_PLANT_2,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         A.REPORT_DATE,                                                   \n";
                    strSQL_Prod = strSQL_Prod + "                                         A.OPER_OUT_QTY1 AS PROD_QTY                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                    FROM UNIERPSEMI.WIPHST@CCUBE A, MIGHTY.PROC_TYPE_INFO@CCUBE B         \n";
                    strSQL_Prod = strSQL_Prod + "                                   WHERE     A.PLANT = 'CCUBEDIGITAL'                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.PLANT = B.PLANT                                            \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.CREATE_CODE = B.PROC_TYPE                                  \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REWORK = 'N'                                               \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REPORT_DATE >= '" + S_date.ToString("yyyyMMdd") + "'       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.REPORT_DATE <= '" + E_date.ToString("yyyyMMdd") + "'       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND (A.ACCOUNT_CODE IS NULL                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                              OR A.ACCOUNT_CODE IN                                        \n";
                    strSQL_Prod = strSQL_Prod + "                                                    (SELECT SCD.SYSCODE_NAME                              \n";
                    strSQL_Prod = strSQL_Prod + "                                                       FROM MIGHTY.SYSCODEDATA@CCUBE SCD                  \n";
                    strSQL_Prod = strSQL_Prod + "                                                      WHERE SCD.PLANT = A.PLANT                           \n";
                    strSQL_Prod = strSQL_Prod + "                                                            AND SCD.SYSTABLE_NAME =                       \n";
                    strSQL_Prod = strSQL_Prod + "                                                                   'SCRAP_REASON'                         \n";
                    strSQL_Prod = strSQL_Prod + "                                                            AND SCD.SYSCODE_GROUP = 'CHARGE'))            \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.LOT_SUB_TYPE = 'NONE'                                      \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.CREATE_CODE <> 'INK'                                       \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.OPERATION <> A.TO_OPERATION                                \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.OPERATION = B.OUT_OPER                                     \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND B.SUB_PLANT_1 IN ('RCP-FINAL', 'RCP-PANEL')                  \n";
                    strSQL_Prod = strSQL_Prod + "                                         AND A.PART_TYPE IN ('P', 'D')                                    \n";
                    strSQL_Prod = strSQL_Prod + "             )                                                                                            \n";
                    strSQL_Prod = strSQL_Prod + "  GROUP BY REPORT_DATE                                                                                    \n";
                    break;
            }
            return strSQL_Prod;
        }


        #endregion

        #region "Event"
        protected void InitTxtWorkDate(object sender, EventArgs e)
        {
            TextBox txtTemp = (TextBox)sender;
            txtTemp.Text = DateTime.Now.ToString("yyyy-MM");
        }

        //protected void Page_PreInit(Object sender, EventArgs e)
        //{
        //    if (Request.QueryString["Menu"] == "false") this.MasterPageFile = "../Site2.Master";
        //}

        protected void Page_Load(object sender, EventArgs e)
        {
            SetDefaultValue();
            drp_Prod_TextChanged();

            if (!IsPostBack)
            {
                Reprt_mcc_Reset(sender, e);
                WebSiteCount();
            }

            if (Request.QueryString["Menu"] == "false")
            {
                //SiteMapPath1.Visible = false;
                Table1.Visible = true;
            }
        }

        protected void query_Click(object sender, EventArgs e)
        {
            string strSQL = "", strCrntDT = "", strDeptSemi = "", strDeptDisp = "", strAreaID = "", strOperID_SAWON = "", strAreaID_SAWON = "", strPartID_SAWON = "";
            string strOperID = "", strPartID = "", strOperID_Disp = "", strAreaID_Disp = "", strOperID_Disp1 = "", strAreaID_Disp1 = "", strSQL_Prod = "";
            DateTime To_date = new DateTime(), S_date = new DateTime(), E_date = new DateTime();

            if (GetDefaultValue(ref strCrntDT, ref strDeptSemi, ref strDeptDisp, ref strAreaID, ref strOperID, ref strPartID, ref strAreaID_Disp,
                                ref strOperID_Disp, ref strAreaID_SAWON, ref strOperID_SAWON, ref strPartID_SAWON, ref strAreaID_Disp1,
                                ref strOperID_Disp1, ref To_date, ref E_date, ref S_date) == false) return;

            strSQL_Prod = MakeQuery_Prod(strSQL_Prod, S_date, E_date);

            if ((strSQL = MakeQuery(strCrntDT, strDeptSemi, strDeptDisp, strAreaID, strOperID, strPartID, strAreaID_SAWON, strOperID_SAWON,
                                    strPartID_SAWON, strAreaID_Disp, strOperID_Disp, strAreaID_Disp1, strOperID_Disp1, strSQL_Prod, To_date, E_date, S_date)) == "") return;



            ERPAppAddition.ERPAddition.INSA.DailyPaySum.CommuteModule dtHM2 = new ERPAppAddition.ERPAddition.INSA.DailyPaySum.CommuteModule();
            GF.CreateReport(dtHM2, 2, strSQL, ReportViewer1, "INSA.DailyPaySumTrend.DailyPaySumTrend.rdlc", "dsDailyPaySumTrend", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());


            //if (strSQL_Prod == "")
            //{
            //    CommuteModule dtHM2 = new CommuteModule();
            //    //GF.CreateReport(dtHM2, 2, strSQL, ReportViewer1, "HM.DailyPaySumTrend.rdlc", "dsDailyPaySumTrend", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());
            //    GF.CreateReport(dtHM2, 2, strSQL, ReportViewer1, "HM.DailyPaySumTrend.rdlc", "dsDailyPaySumTrend", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());
            //}
            //else
            //{
            //    CommuteModule dtHM2 = new CommuteModule();
            //    //CommuteModule dtHM3 = new CommuteModule();
            //    //GF.CreateReport_DataSet2(dtHM2, 0, 1, strSQL, strSQL_Prod, ReportViewer1, "HM.DailyPaySumTrend_Pro.rdlc", "dsDailyPaySumTrend", "dsDailyPaySumTrendProd", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());
            //    GF.CreateReport(dtHM2, 1, strSQL, ReportViewer1, "HM.DailyPaySumTrend_Pro.rdlc", "dsDailyPaySumTrendProd", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());
            //}
        }

        protected void Reprt_mcc_Reset(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            mcc_dr_Area.ClearSQLText();
            mcc_dr_Part.ClearSQLText();
            mcc_dr_Oprgrp.ClearSQLText();
            drp_Prod.ClearSelection();  // 151114 변수초기화 추가
            drp_Prod_sub.ClearSelection();  // 151114 변수초기화 추가 
        }

        protected void Reprt_Reset(object sender, EventArgs e)
        {
            ReportViewer1.Reset();
        }

        protected void drp_Prod_TextChanged()
        {
            if (dr_dept.SelectedValue == "Semi")
            {
                drp_Prod.Visible = true;
                td_id.Visible = true;

                if (drp_Prod.SelectedValue == "B" || drp_Prod.SelectedValue == "C" || drp_Prod.SelectedValue == "E"
                    || drp_Prod.SelectedValue == "F")
                {
                    drp_Prod_sub.Visible = true;
                }
                else
                {
                    drp_Prod_sub.Visible = false;
                }
            }
            else
            {
                drp_Prod.Visible = false;
                td_id.Visible = false;
                drp_Prod_sub.Visible = false;
            }
        #endregion
        }
    }
}