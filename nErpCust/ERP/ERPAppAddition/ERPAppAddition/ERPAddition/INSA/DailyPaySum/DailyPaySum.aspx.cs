using System.Data;
using System;
using System.Web.UI.WebControls;
using SRL.UserControls;
using System.Text;
using System.Web;

namespace ERPAppAddition.ERPAddition.INSA.DailyPaySum
{
    public partial class DailyPaySum : System.Web.UI.Page
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
                    GF.SetCostIDCtrl(GV.gOraCon, mcc_dr_Cost, Depart); // 150816 신규추가
                    GF.SetDepartIDCtrl(GV.gOraCon, mcc_dr_Depart, Depart); //150816 신규추가
                }
                else if (dr_dept.SelectedItem.Text == "Display")
                {
                    Depart = "Display";

                    GV.gOraCon.Open();
                    mcc_dr_Part.ClearAll();
                    mcc_dr_Cost.ClearAll(); //150816 신규추가
                    mcc_dr_Depart.ClearAll(); //150816 신규추가
                    GF.SetOprGroupIDCtrl(GV.gOraCon, mcc_dr_Oprgrp, Depart);
                    GF.SetAreaIDCtrl(GV.gOraCon, mcc_dr_Area, Depart);
                }
                else
                {
                    mcc_dr_Part.ClearAll();
                    mcc_dr_Oprgrp.ClearAll();
                    mcc_dr_Area.ClearAll();
                    mcc_dr_Cost.ClearAll(); //150816 신규추가
                    mcc_dr_Depart.ClearAll(); //150816 신규추가
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

        private string MakeQuery(string strCrntDT, string strdeptID)
        {
            string strSQL;

            strSQL = "SELECT * FROM (  \n";
            strSQL = strSQL + " WITH DEFUALT AS ( \n";
            strSQL = strSQL + " SELECT                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "                      F.AREA                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "                     ,F.AREA AS PART                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "                     ,F.OPER_GROUP                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "                     ,F.VENDOR                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                     ,F.WORK_GROUP                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "                     ,ROUND( (SUM(F.NWRK_TIME - F.LATE_EARLYLEAVE)) + SUM(F.OW_TIME) + SUM(F.HOLWRK_TIME) + SUM(F.HOLWRK_OW_TIME), 1) AS TIME_SUM -- 근태정보합계                                                                                                    \n";
            strSQL = strSQL + "                     ,SUM(F.NWRK_TIME - F.LATE_EARLYLEAVE) AS NWRK_TIME -- 평일근무 (평일정상시간 - 지각조퇴 시간)                                                                                                                                                   \n";
            strSQL = strSQL + "                     ,SUM(F.OW_TIME) AS OW_TIME --평일연장                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                     ,SUM(F.NIGHT_TIME) AS NIGHT_TIME -- 야간                                                                                                                                                                                                        \n";
            strSQL = strSQL + "                     ,SUM(F.HOLWRK_TIME) AS HOLWRK_TIME --휴일                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                     ,SUM(F.HOLWRK_OW_TIME) AS HOLWRK_OW_TIME -- 휴일연장                                                                                                                                                                                            \n";
            strSQL = strSQL + "                     ,ROUND( SUM(F.BASE_PAY)                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "                     + SUM(F.OW_TIME_PAY)                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "                     + SUM(F.HOLWRK_OW_TIME_PAY)                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "                     + SUM(F.HOLWRK_TIME_PAY)                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "                     + SUM(F.NIGHT_TIME_PAY)                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "                     + SUM(F.ALLOWANCE_2_PAY)                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "                     + ( (SUM(F.BASE_PAY) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) /12 )                                                                                          \n";
            strSQL = strSQL + "                     + ( (SUM(F.BASE_PAY) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) * 0.09979)                                                                                     \n";
            strSQL = strSQL + "                     + SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD')) )                                                                                                                                                                                 \n";
            strSQL = strSQL + "                     + SUM(F.DATA_1_PANGONG_PAY + F.DATA_1_LEVEL_PAY) ,0) AS PAY_SUM -- 급여정보 합계                                                                                                                                                                \n";
            strSQL = strSQL + "                     ,SUM(ROUND(F.BASE_PAY, 0)) AS BASE_PAY -- 기본급                                                                                                                                                                                                \n";
            strSQL = strSQL + "                     ,SUM(F.OW_TIME_PAY) AS OW_TIME_PAY -- 평일연장                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                     ,SUM(F.HOLWRK_OW_TIME_PAY) AS HOLWRK_OW_TIME_PAY -- 휴일연장                                                                                                                                                                                    \n";
            strSQL = strSQL + "                     ,SUM(F.HOLWRK_TIME_PAY) AS HOLWRK_TIME_PAY -- 휴일근무                                                                                                                                                                                          \n";
            strSQL = strSQL + "                     ,SUM(F.NIGHT_TIME_PAY) AS NIGHT_TIME_PAY -- 야간                                                                                                                                                                                                \n";
            strSQL = strSQL + "                     ,ROUND(SUM(F.ALLOWANCE_2_PAY),0) AS ALLOWANCE_2_PAY --상여금                                                                                                                                                                                    \n";
            strSQL = strSQL + "                     ,ROUND((SUM(ROUND(F.BASE_PAY, 0)) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) / 12, 0) AS RESIGN_PAY                                                            \n";
            strSQL = strSQL + "                     ,ROUND(((SUM(ROUND(F.BASE_PAY, 0)) + SUM(F.OW_TIME_PAY) + SUM(F.HOLWRK_OW_TIME_PAY) + SUM(F.HOLWRK_TIME_PAY) + SUM(F.NIGHT_TIME_PAY) + SUM(F.ALLOWANCE_2_PAY)) * 0.09979), 0) AS INSURANCE_PAY                                                  \n";
            strSQL = strSQL + "                     ,ROUND(SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD'))),0) AS MAINT_PAY -- 관리비                                                                                                                                                   \n";
            strSQL = strSQL + "                     ,ROUND(SUM(F.DATA_1_PANGONG_PAY + F.DATA_1_LEVEL_PAY),0) AS ETC_PAY -- 기타급여                                                                                                                                                                 \n";
            strSQL = strSQL + "             FROM (                                                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                                        SELECT DISTINCT TIME_DATA.*     -- 하위쿼리의 기본 인적사항 및 시간                                                                                                                                                          \n";
            strSQL = strSQL + "                                                 , C.HOURLY_WAGE AS HOURLY_WAGE_PAY -- 개인시급                                                                                                                                                                      \n";
            strSQL = strSQL + "                                                 , CASE WHEN TIME_DATA.WORK_GROUP  IN ('제조1파트','제조2파트') THEN C.HOURLY_WAGE * 209 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') ) * WRK_DAY                                                                \n";
            strSQL = strSQL + "                                                            ELSE (C.MONTHLY_WAGE + C.ALLOWANCE_7)  / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') )                                                                                                                \n";
            strSQL = strSQL + "                                                    END AS BASE_PAY   -- 기본급 (시급 * 209 * 해당 월 MAX일수 * 실제 근무일수) / 디스플레이는 실제 휴                                                                                                                                        \n";
            strSQL = strSQL + "                                                 , (TIME_DATA.OW_TIME * C.HOURLY_WAGE * 1.5) AS OW_TIME_PAY   -- 평일연장수당 ( 시급 * 1.5 * 평일연장근무시간 )                                                                                                      \n";
            strSQL = strSQL + "                                                 , (TIME_DATA.HOLWRK_OW_TIME * C.HOURLY_WAGE * 2) AS HOLWRK_OW_TIME_PAY   -- 주말연장수당 ( 시급 * 2 * 주말연장근무시간)                                                                                             \n";
            strSQL = strSQL + "                                                 , CASE WHEN  TIME_DATA.WORK_GROUP IN ('제조1파트', '제조2파트') THEN TIME_DATA.HOLWRK_TIME * C.HOURLY_WAGE * 1.5                                                                                                    \n";
            strSQL = strSQL + "                                                            ELSE (TIME_DATA.HOLWRK_TIME / 4 ) * 30000                                                                                                                                                                \n";
            strSQL = strSQL + "                                                   END AS HOLWRK_TIME_PAY    -- 특근수당 ( 시급 * 1.5 * 주말근무시간 )                                                                                                                                               \n";
            strSQL = strSQL + "                                                 , (TIME_DATA.NIGHT_TIME * C.HOURLY_WAGE * 0.5) AS NIGHT_TIME_PAY   -- 야간수당 ( 야간 근무시간 * 시급 * 0.5 )                                                                                                       \n";
            strSQL = strSQL + "                                                 , TIME_DATA.LATE_EARLYLEAVE * C.HOURLY_WAGE AS LATE_EARLYLEAVE_PAY            -- 지각조퇴 차감 ( 시급 * 지각조퇴시간 )                                                                                              \n";
            strSQL = strSQL + "                                                 ,  CASE WHEN TIME_DATA.WORK_GROUP  IN ('제조1파트','제조2파트') THEN (C.HOURLY_WAGE * 209) / 12 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') ) * WRK_DAY                                                        \n";
            strSQL = strSQL + "                                                             ELSE 0                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                                                    END AS ALLOWANCE_2_PAY          -- 상여금 (개인별 입력수당)                                                                                                                                                      \n";
            strSQL = strSQL + "                                                 ,  TIME_DATA.DATA_1_PANGONG /  (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD'))  AS DATA_1_PANGONG_PAY       -- 판공비 (기준정보 고정수당 )                                                                           \n";
            strSQL = strSQL + "                                                 ,  TIME_DATA.DATA_1_LEVEL /  (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD')) AS DATA_1_LEVEL_PAY   --인증수당 ( 기준정보 고정수당 )                                                                                  \n";
            strSQL = strSQL + "                                         FROM (                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "                                                 SELECT DISTINCT                                                                                                                                                                                                     \n";
            strSQL = strSQL + "                                                         B.VENDOR   -- 업체                                                                                                                                                                                          \n";
            strSQL = strSQL + "                                                         , CASE B.WORK_GROUP when '1' THEN '제조1파트'                                                                                                                                                               \n";
            strSQL = strSQL + "                                                                                      when '2' THEN '제조2파트'                                                                                                                                                      \n";
            strSQL = strSQL + "                                                                                      when 'N' THEN '주간'                                                                                                                                                           \n";
            strSQL = strSQL + "                                                            END AS WORK_GROUP   -- 근무조                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                         , B.OPER_GROUP   -- 공정                                                                                                                                                                                    \n";
            strSQL = strSQL + "                                                         , B.AREA -- 층수                                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                         , B. EMP_ID -- 사번                                                                                                                                                                                         \n";
            strSQL = strSQL + "                                                         , CASE WHEN (B.NW_TIME = 0 OR B.ABS_TYPE IN ('무급', '병가')) THEN 0                                                                                                                                        \n";
            strSQL = strSQL + "                                                                ELSE 1                                                                                                                                                                                               \n";
            strSQL = strSQL + "                                                           END AS WRK_DAY                                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                         , NVL(SUM(B.LATE_TIME + B.EALV_TIME), 0) AS LATE_EARLYLEAVE    -- 지각조퇴(시간)                                                                                                                            \n";
            strSQL = strSQL + "                                                         , CASE WHEN  B.WORK_GROUP IN ('1', '2') THEN NVL(SUM(DECODE(B.HOLWRK_FLAG,'', B.OW_TIME, NULL, B.OW_TIME)), 0)                                                                                              \n";
            strSQL = strSQL + "                                                                    ELSE 0                                                                                                                                                                                           \n";
            strSQL = strSQL + "                                                            END AS OW_TIME   -- 평일연장(시간)                               --                                                                                                                                      \n";
            strSQL = strSQL + "                                                         , CASE WHEN B.WORK_GROUP IN ('1', '2') THEN  NVL(SUM(DECODE(B.HOLWRK_FLAG,'Y', B.OW_TIME,0)), 0)                                                                                                            \n";
            strSQL = strSQL + "                                                                    ELSE 0                                                                                                                                                                                           \n";
            strSQL = strSQL + "                                                            END AS HOLWRK_OW_TIME   -- 휴일연장(시간)                                                                                                                                                                \n";
            //strSQL = strSQL + "                                                         , NVL(SUM(DECODE(B.HOLWRK_FLAG,'Y', B.NW_TIME,0)), 0) AS HOLWRK_TIME   -- 특근(시간)                                                                                                                      \n";
            strSQL = strSQL + "                                                        , SUM(CASE WHEN B.HOLWRK_FLAG = 'Y' AND (B.ABS_TYPE IS NULL OR B.ABS_TYPE != '휴무') THEN B.NW_TIME - B.LATE_TIME                                                                                            \n";
            strSQL = strSQL + "                                                                   ELSE 0                                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                              END) AS HOLWRK_TIME   -- 특근(시간)                                                                                                                                                                    \n";
            strSQL = strSQL + "                                                         , CASE WHEN  B.WORK_GROUP IN ('1', '2') THEN NVL(SUM(B.NIGHT_TIME), 0)                                                                                                                                      \n";
            strSQL = strSQL + "                                                                    ELSE 0                                                                                                                                                                                           \n";
            strSQL = strSQL + "                                                            END AS NIGHT_TIME   -- 야간(시간)                                                                                                                                                                        \n";
            strSQL = strSQL + "                                                         , DECODE(D.DATA_1, '', 0,  NULL, 0, D.DATA_1) AS DATA_1_PANGONG   -- 판공비(직급수당)                                                                                                                       \n";
            strSQL = strSQL + "                                                         , DECODE(E.DATA_1, '', 0,  NULL, 0, E.DATA_1) AS DATA_1_LEVEL   --인센티브(인증수당)                                                                                                                        \n";
            strSQL = strSQL + "                                                         , NVL(SUM(CASE WHEN B.ABS_TYPE IN ('휴무', '공가', '연차') OR B.HOLWRK_FLAG = 'Y' THEN 0                                                                                                                    \n";
            strSQL = strSQL + "                                                               ELSE B.NW_TIME                                                                                                                                                                                        \n";
            strSQL = strSQL + "                                                               END),0) AS NWRK_TIME   -- 평일(시간))                                                                                                                                                                 \n";
            //strSQL = strSQL + "                                                         , NVL(SUM(DECODE(B.HOLWRK_FLAG, NULL, B.NW_TIME)), 0) AS NWRK_TIME   -- 평일(시간)                                                                                                                         \n";
            strSQL = strSQL + "                                                         , ABS_TYPE                                                                                                                                                                                                  \n";
            strSQL = strSQL + "                                                 FROM NHRMWRKDEF A INNER JOIN NHRMWRKIOT B ON A.EMP_ID = B.EMP_ID AND B.NW_TIME >= 1                                                                                                                                 \n";
            strSQL = strSQL + "                                                                                   LEFT OUTER JOIN MGCMTBLDAT D ON A.POSITION = D.KEY_2 AND D.FACTORY = 'DISPLAY' AND D.TABLE_NAME = 'ALLOWANCE'                                                                     \n";
            strSQL = strSQL + "                                                                                   LEFT OUTER JOIN MGCMTBLDAT E ON A.CERT_LEVEL = E.KEY_2 AND E.FACTORY = 'DISPLAY' AND E.TABLE_NAME = 'ALLOWANCE'                                                                   \n";
            strSQL = strSQL + "                                                 WHERE B.WORK_DATE = '" + strCrntDT + "' /*이쪽에 각종 조건문을 넣자*/                                                                                                                                               \n";
            strSQL = strSQL + "                                                       AND A.RESIGN_TYPE <> '입사취소'    /* 퇴사사유 중 재직 3일 미만 (입사취소자) 데이터는 출력하지 않도록 수정 [2015.11.11, ahncj, 생산팀 윤미화 대리 요청] */                                                               \n";
            strSQL = strSQL + "                                                 GROUP BY B.VENDOR, B.OPER_GROUP, B.OPER_GROUP, B.AREA, B.WORK_GROUP, D.DATA_1, E.DATA_1, B.EMP_ID, B.NW_TIME, B.ABS_TYPE                                                                                            \n";
            strSQL = strSQL + "                                                 )TIME_DATA                                                                                                                                                                                                          \n";
            strSQL = strSQL + "                                    INNER JOIN (                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "                                       SELECT B.*                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                                                 FROM (                                                                                                                                                                                                              \n";
            strSQL = strSQL + "                                                       SELECT MAX(WORK_DATE) AS WORK_DATE, EMP_ID                                                                                                                                                                    \n";
            strSQL = strSQL + "                                                        FROM NHRMWRKPAY                                                                                                                                                                                              \n";
            strSQL = strSQL + "                                                         WHERE WORK_DATE <= '" + strCrntDT + "'                                                                                                                                                                      \n";
            strSQL = strSQL + "                                                        GROUP BY EMP_ID) A, NHRMWRKPAY B                                                                                                                                                                             \n";
            strSQL = strSQL + "                                                WHERE 1=1                                                                                                                                                                                                            \n";
            strSQL = strSQL + "                                                 AND A.WORK_DATE = B.WORK_DATE                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                                 AND A.EMP_ID = B.EMP_ID) C ON TIME_DATA.EMP_ID = C.EMP_ID                                                                                                                                                           \n";
            strSQL = strSQL + "             ) F                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "             GROUP BY                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "              F.OPER_GROUP                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "             ,F.AREA                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "             ,F.VENDOR                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "            ,F.WORK_GROUP                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "             ORDER BY F.AREA, F.OPER_GROUP, F.WORK_GROUP                                                                                                                                                                                                             \n";
            strSQL = strSQL + " )                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + " ,                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + " HUMAN_TOTAL AS ( -- 재적인원                                                                                                                                                                                                                                        \n";
            strSQL = strSQL + "     SELECT                                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "           AREA                                                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "         , OPER_GROUP                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "         , VENDOR                                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "         , WORK_GROUP                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "         , WORK_DATE                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "         , COUNT(*) AS CTT                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "     FROM NHRMWRKIOT                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "       WHERE 1=1                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "       AND WORK_DATE = '" + strCrntDT + "'                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "     GROUP BY WORK_DATE, AREA, OPER_GROUP, WORK_DATE, VENDOR, WORK_GROUP                                                                                                                                                                                             \n";
            strSQL = strSQL + " )                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + " ,                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + " HUMAN_SUM AS ( -- 출근안한 인원                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "     SELECT                                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "           AREA                                                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "         , OPER_GROUP                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "         , VENDOR                                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "         , WORK_GROUP                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "         , WORK_DATE                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "         , COUNT(*) AS CT                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "     FROM NHRMWRKIOT                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "       WHERE 1=1                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "       AND WORK_DATE = '" + strCrntDT + "'                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "       AND ABS_TYPE NOT IN  ('정상', '지각', '반차', '비상', '조퇴', '반공')                                                                                                                                                                                         \n";
            strSQL = strSQL + "     GROUP BY  WORK_DATE, AREA, OPER_GROUP, WORK_DATE, VENDOR, WORK_GROUP                                                                                                                                                                                            \n";
            strSQL = strSQL + " )                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "  ,                                                                                                                                                                                                                                                                  \n";
            strSQL = strSQL + " HUMAN AS ( -- 재적인원 - 근태인원 (재적인원 - 출근안한 인원)                                                                                                                                                                                                        \n";
            strSQL = strSQL + "     SELECT                                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "         HUMAN_TOTAL.AREA                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "       , HUMAN_TOTAL.OPER_GROUP                                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "       , HUMAN_TOTAL.VENDOR                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "       , HUMAN_TOTAL.WORK_DATE AS WORK_DATE                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "       , CASE HUMAN_TOTAL.WORK_GROUP when '1' THEN '제조1파트'                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                     when '2' THEN '제조2파트'                                                                                                                                                                                                       \n";
            strSQL = strSQL + "                                     when 'N' THEN '주간'                                                                                                                                                                                                            \n";
            strSQL = strSQL + "         END AS WORK_GROUP   -- 근무조                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "       , HUMAN_TOTAL.CTT AS CTT                                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "       , (HUMAN_TOTAL.CTT - NVL(HUMAN_SUM.CT, 0)) AS CT                                                                                                                                                                                                              \n";
            strSQL = strSQL + "       , ROUND(( (HUMAN_TOTAL.CTT - NVL(HUMAN_SUM.CT, 0))  / HUMAN_TOTAL.CTT) * 100, 0) || '%' AS RATIO                                                                                                                                                              \n";
            strSQL = strSQL + "     FROM HUMAN_TOTAL, HUMAN_SUM                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "       WHERE 1=1                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "       AND HUMAN_TOTAL.WORK_DATE = HUMAN_SUM.WORK_DATE(+)                                                                                                                                                                                                            \n";
            strSQL = strSQL + "       AND HUMAN_TOTAL.OPER_GROUP = HUMAN_SUM.OPER_GROUP(+)                                                                                                                                                                                                          \n";
            strSQL = strSQL + "       AND HUMAN_TOTAL.AREA = HUMAN_SUM.AREA (+)                                                                                                                                                                                                                     \n";
            strSQL = strSQL + "       AND HUMAN_TOTAL.VENDOR = HUMAN_SUM.VENDOR (+)                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "       AND HUMAN_TOTAL.WORK_GROUP = HUMAN_SUM.WORK_GROUP (+)                                                                                                                                                                                                         \n";
            strSQL = strSQL + " )                                                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "                                                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + " select *                                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + " from(                                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + " SELECT                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "       TO_CHAR(Z.AREA) AS AREA                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "     , TO_CHAR(Z.OPER_GROUP) AS OPER_GROUP                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "     , TO_CHAR(Z.VENDOR) AS VENDOR                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "     , TO_CHAR(Z.WORK_GROUP) AS WORK_GROUP                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "     , TO_CHAR(Z.PART) AS PART                                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "     , '-' AS S_COST                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "     , '-' AS S_DEPT                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "     , TO_NUMBER(Y.CTT) AS FULL_HUMAN                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "     , TO_NUMBER(Y.CT) AS WORK_HUMAN                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "     , TO_CHAR(Y.RATIO) AS PROPORTION                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.TIME_SUM) AS TIME_SUM                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.NWRK_TIME) AS NWRK_TIME                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.OW_TIME) AS OW_TIME                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.NIGHT_TIME) AS NIGHT_TIME                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.HOLWRK_TIME) AS HOLWRK_TIME                                                                                                                                                                                                                       \n";
            strSQL = strSQL + "     , TO_NUMBER(Z.HOLWRK_OW_TIME) AS HOLWRK_OW_TIME                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.PAY_SUM),0) AS PAY_SUM                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.BASE_PAY),0) AS BASE_PAY                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.OW_TIME_PAY),0) AS OW_TIME_PAY                                                                                                                                                                                                              \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.HOLWRK_OW_TIME_PAY),0) AS HOLWRK_OW_TIME_PAY                                                                                                                                                                                                \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.HOLWRK_TIME_PAY),0) AS HOLWRK_TIME_PAY                                                                                                                                                                                                      \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.NIGHT_TIME_PAY),0) AS NIGHT_TIME_PAY                                                                                                                                                                                                        \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.ALLOWANCE_2_PAY),0) AS ALLOWANCE_2_PAY                                                                                                                                                                                                      \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.RESIGN_PAY),0) AS RESIGN_PAY                                                                                                                                                                                                                \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.INSURANCE_PAY),0) AS INSURANCE_PAY                                                                                                                                                                                                          \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.MAINT_PAY),0) AS MAINT_PAY                                                                                                                                                                                                                  \n";
            strSQL = strSQL + "     , ROUND(TO_NUMBER(Z.ETC_PAY),0) AS ETC_PAY                                                                                                                                                                                                                      \n";

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                strSQL = strSQL + "     , 'Display' AS DEPT                                                                                                                                                                                                                                     \n";
            }

            strSQL = strSQL + " FROM DEFUALT Z, HUMAN Y                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + " WHERE 1=1                                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + " AND Y.AREA = Z.AREA                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + " AND Y.OPER_GROUP = Z.OPER_GROUP                                                                                                                                                                                                                                     \n";
            strSQL = strSQL + " AND Y.VENDOR = Z.VENDOR  AND Y.WORK_GROUP = Z.WORK_GROUP)                                                                                                                                                                                                           \n";
            strSQL = strSQL + "UNION ALL                                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "SELECT                                                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "TO_CHAR(S_GROUP) AS AREA                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + ",TO_CHAR(S_BAY) AS OPER_GROUP                                                                                                                                                                                                                                        \n";
            strSQL = strSQL + ",TO_CHAR(S_OS) AS VENDOR                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + ",TO_CHAR(S_JO) AS WORK_GROUP                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + ",ASSIGN_PART AS PART                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + ",S_COST                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + ",S_DEPT                                                                                                                                                                                                                                                              \n";
            strSQL = strSQL + ",TO_NUMBER(COUNT(SAWON_NO)) AS FULL_HUMAN                                                                                                                                                                                                                            \n";
            strSQL = strSQL + ",TO_NUMBER(SUM(W_CNT)) AS WORK_HUMAN                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + ",TO_CHAR(ROUND((SUM(W_CNT)/COUNT(SAWON_NO)*100),1) || '%')  AS  PROPORTION                                                                                                                                                                                           \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(NW_TIME) + SUM(OW_TIME) + SUM(HOLWRK_TIME) + SUM(HOLWRK_OW_TIME),1)) AS TIME_SUM -- 근태정보합계                                                                                                                                          \n";
            strSQL = strSQL + "      ,TO_NUMBER((SUM(NW_TIME)))AS NWRK_TIME -- 평일근무 (평일정상시간 - 지각조퇴 시간)                                                                                                                                                                              \n";
            strSQL = strSQL + "      ,TO_NUMBER(SUM(OW_TIME)) AS OW_TIME --평일연장                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "      ,TO_NUMBER(SUM(NIGHT_TIME)) AS NIGHT_TIME -- 야간                                                                                                                                                                                                              \n";
            strSQL = strSQL + "      ,TO_NUMBER(SUM(HOLWRK_TIME)) AS HOLWRK_TIME --휴일                                                                                                                                                                                                             \n";
            strSQL = strSQL + "      ,TO_NUMBER(SUM(HOLWRK_OW_TIME)) AS HOLWRK_OW_TIME -- 휴일연장                                                                                                                                                                                                  \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND( SUM(ROUND(BASE_PAY, 0))                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "      + SUM(OW_TIME_PAY)                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "      + SUM(HOLWRK_OW_TIME_PAY)                                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "      + SUM(HOLWRK_TIME_PAY)                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "      + SUM(NIGHT_TIME_PAY)                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "      + SUM(ALLOWANCE_2_PAY)                                                                                                                                                                                                                                         \n";
            strSQL = strSQL + "      + ((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) / 12)                                                                                                            \n";
            strSQL = strSQL + "      + ((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) * 0.09979)                                                                                                       \n";
            strSQL = strSQL + "      + SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD')))                                                                                                                                                                                                 \n";
            strSQL = strSQL + "      + SUM (DATA_1_PANGONG_PAY )                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "      )) AS PAY_SUM -- 급여정보 합계                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "      ,TO_NUMBER(SUM(ROUND(BASE_PAY, 0))) AS BASE_PAY -- 기본급                                                                                                                                                                                                      \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(OW_TIME_PAY), 0)) AS OW_TIME_PAY -- 평일연장                                                                                                                                                                                              \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(HOLWRK_OW_TIME_PAY), 0)) AS HOLWRK_OW_TIME_PAY -- 휴일연장                                                                                                                                                                                \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(HOLWRK_TIME_PAY) , 0))AS HOLWRK_TIME_PAY -- 휴일근무                                                                                                                                                                                      \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(NIGHT_TIME_PAY), 0)) AS NIGHT_TIME_PAY -- 야간                                                                                                                                                                                            \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(ALLOWANCE_2_PAY), 0)) AS ALLOWANCE_2_PAY --상여금                                                                                                                                                                                         \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) / 12, 0)) AS RESIGN_PAY                                                                            \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND( ((SUM(ROUND(BASE_PAY, 0)) + SUM(OW_TIME_PAY) + SUM(HOLWRK_OW_TIME_PAY) + SUM(HOLWRK_TIME_PAY) + SUM(NIGHT_TIME_PAY) + SUM(ALLOWANCE_2_PAY)) * 0.09979), 0)) AS INSURANCE_PAY                                                                 \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(115000 / (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD'))),0)) AS MAINT_PAY -- 관리비                                                                                                                                                       \n";
            strSQL = strSQL + "      ,TO_NUMBER(ROUND(SUM(DATA_1_PANGONG_PAY), 0)) AS ETC_PAY -- 기타급여                                                                                                                                                                                           \n";

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                strSQL = strSQL + "   , 'Semi' AS DEPT \n";
            }

            strSQL = strSQL + "FROM                                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "   (                                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "    SELECT A.SAWON_NO                                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "          ,B.ASSIGN_GROUP,B.ASSIGN_PART,B.ASSIGN_BAY,B.SAWON_OUTSOURCING,B.ASSIGN_JOIRUM                                                                                                                                                                             \n";
            strSQL = strSQL + "          ,S_GROUP,S_GROUP_SEQ,S_JO,S_BAY,S_BAY_SEQ,S_OS,S_COST, S_DEPT                                                                                                                                                                                              \n";
            strSQL = strSQL + "          ,CASE WHEN NW_TIME IS NOT NULL OR NW_TIME > 0 THEN 1 ELSE 0 END W_CNT                                                                                                                                                                                      \n";
            strSQL = strSQL + "     ,NVL(DECODE(HOLIDAY_WORK,NULL,NW_TIME),0)NW_TIME                                                                                                                                                                                                                \n";
            strSQL = strSQL + "          ,NVL(DECODE(HOLIDAY_WORK,'O',NW_TIME),0)HOLWRK_TIME                                                                                                                                                                                                        \n";
            strSQL = strSQL + "          ,NVL(DECODE(HOLIDAY_WORK,NULL,OW_TIME),0)OW_TIME                                                                                                                                                                                                           \n";
            strSQL = strSQL + "          ,NVL(DECODE(HOLIDAY_WORK,'O',OW_TIME),0)HOLWRK_OW_TIME                                                                                                                                                                                                     \n";
            strSQL = strSQL + "     ,NVL(NIGHT_TIME,0)NIGHT_TIME                                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "          ,VACATION_KIND                                                                                                                                                                                                                                             \n";
            strSQL = strSQL + "          ,NVL(LATE_TIME,0)LATE_TIME                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "          ,NVL(EARLY_TIME,0)EARLY_TIME                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "          ,D.SAWON_PAY_GBN                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "          ,TIME_PAY                                                                                                                                                                                                                                                  \n";
            //strSQL = strSQL + "          ,NVL(CASE WHEN SAWON_PAY_GBN ='시급' THEN TIME_PAY * 209 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') )                                                                                                                                                \n";
            //strSQL = strSQL + "                    ELSE (MONTH_PAY + EXTRA_PAY) / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') )                                                                                                                                                                  \n";
            //strSQL = strSQL + "               END,0) AS BASE_PAY   -- 기본급 (시급 * 209 * 해당 월 MAX일수 )                                                                                                                                                                                        \n";
            strSQL = strSQL + "          ,CASE WHEN VACATION_KIND IN ('출산휴가', '병가', '무급', '생휴') THEN 0 --실제 무급이 발생되는 카테고리는 기본급을 0원 처리 (좌측 4개 내용 외 기본급 발생)                                                                                                                   \n";
            strSQL = strSQL + "                                                       ELSE ( NVL(CASE WHEN SAWON_PAY_GBN ='시급' THEN TIME_PAY * 209 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') )                                                                                                   \n";
            strSQL = strSQL + "                                                                       ELSE (MONTH_PAY + EXTRA_PAY) / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') )                                                                                                                    \n";
            strSQL = strSQL + "                                                              END,0) )                                                                                                                                                                                             \n";
            strSQL = strSQL + "           END AS BASE_PAY   -- 기본급 (시급 * 209 * 해당 월 MAX일수 )                                                                                                                                                                                                \n";
            strSQL = strSQL + "          ,NVL(DECODE(HOLIDAY_WORK,NULL, NVL(OW_TIME,0) * TIME_PAY * 1.5),0) OW_TIME_PAY  -- 평일연장수당 ( 시급 * 1.5 * 평일연장근무시간 )                                                                                                                          \n";
            strSQL = strSQL + "          ,NVL(CASE WHEN SAWON_PAY_GBN ='시급'  AND HOLIDAY_WORK ='O' THEN NVL(NW_TIME,0) * TIME_PAY * 1.5                                                                                                                                                           \n";
            strSQL = strSQL + "                    WHEN SAWON_PAY_GBN ='월급'  AND HOLIDAY_WORK ='O' THEN (NVL(NW_TIME,0) / 4 ) * 30000                                                                                                                                                             \n";
            strSQL = strSQL + "               END,0) AS HOLWRK_TIME_PAY    -- 특근수당 ( 시급 * 1.5 * 주말근무시간 )                                                                                                                                                                                \n";
            strSQL = strSQL + "          ,NVL(DECODE(HOLIDAY_WORK,'O', NVL(OW_TIME,0) * TIME_PAY * 2),0) HOLWRK_OW_TIME_PAY   -- 주말연장수당 ( 시급 * 2 * 주말연장근무시간)                                                                                                                        \n";
            strSQL = strSQL + "          ,NVL((NVL(NIGHT_TIME,0) * TIME_PAY * 0.5),0) NIGHT_TIME_PAY   -- 야간수당 ( 야간 근무시간 * 시급 * 0.5 )                                                                                                                                                   \n";
            strSQL = strSQL + "          ,NVL((NVL(LATE_TIME,0) + NVL(EARLY_TIME,0)) * TIME_PAY,0) LATE_EARLYLEAVE_PAY            -- 지각조퇴 차감 ( 시급 * 지각조퇴시간 )                                                                                                                          \n";
            strSQL = strSQL + "          ,NVL(DECODE(SAWON_PAY_GBN,'시급', (TIME_PAY * 209 )/12 / ( TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD') ) * (CASE WHEN NW_TIME IS NOT NULL OR NW_TIME > 0 THEN 1 ELSE 0 END)),0) ALLOWANCE_2_PAY  -- 상여금(월급직은 상여없음)                             \n";
            strSQL = strSQL + "          ,NVL((NVL(LEVEL_PAY,0) +NVL( ETC_PAY,0)) / (TO_CHAR(LAST_DAY('" + strCrntDT + "'), 'DD')),0) DATA_1_PANGONG_PAY       -- 판공비 (기준정보 고정수당 )                                                                                                       \n";
            //strSQL = strSQL + "           FROM(SELECT * FROM SAWON_INF@CCUBE WHERE SAWON_JOIN_DATE <= TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD' ) and ( SAWON_QUIT_DATE IS NULL OR SAWON_QUIT_DATE >= TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD' )))A                              \n"; 입사취소자 제외 관련 추가 (2015.11.11, ahncj, 생산팀 윤미화 대리님 요청)
            //strSQL = strSQL + "           FROM(SELECT * FROM SAWON_INF@CCUBE WHERE SAWON_JOIN_DATE <= TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD' ) and ( SAWON_QUIT_DATE IS NULL OR SAWON_QUIT_DATE >= TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD' ) ))A   \n";
            // 3일미만 입사자 제외요청 관련 신규쿼리 적용 (이후 서브쿼리 중 'A' 테이블)
            strSQL = strSQL + "           FROM( SELECT *                                                                                                                                                                                            \n";
            strSQL = strSQL + "                 FROM SAWON_INF@CCUBE WHERE SAWON_JOIN_DATE <= '" + strCrntDT + "' AND (SAWON_QUIT_DATE >= '" + strCrntDT + "' OR SAWON_QUIT_DATE IS NULL)                                                  \n";
            strSQL = strSQL + "                      AND (SAWON_QUIT_GBN <> '입사취소' OR SAWON_QUIT_GBN IS NULL) AND SAWON_DIVISION = '반도체' )A                                                                                                                                                                        \n";
            //strSQL = strSQL + "                  AND SAWON_DIVISION='반도체'                                                                                                                                                                                                                        \n";
            //strSQL = strSQL + "                  AND (CASE WHEN SAWON_QUIT_DATE = TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD') AND SAWON_QUIT_GBN ='입사취소' THEN 'Y'  END) IS NULL) A                                                                                                                                             \n";
            strSQL = strSQL + "              ,(SELECT * FROM SAWON_ASSIGN_INF@CCUBE WHERE DEL_FLAG IS NULL AND ASSIGN_SECT ='1')B                                                                                                                                                                   \n";
            strSQL = strSQL + "              ,(SELECT * FROM SAWON_WORK_TIME_INF@CCUBE WHERE WORK_DATE = TO_CHAR( TO_DATE('" + strCrntDT + "'),'YYYY-MM-DD') )C                                                                                                                                     \n";
            //strSQL = strSQL + "              ,(SELECT * FROM SAWON_PAY_INF@CCUBE WHERE SECT='1')D                                                                                                                                                                                                 \n";
            //인상시급 관련 쿼리 적용 (조회날짜에 따라 인상된 시급이 변경되어 조회되게끔 수정)                                                                                                                                                                                                                                                                                      
            strSQL = strSQL + "               ,(SELECT A.*                                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "                    FROM FTUSER.SAWON_PAY_INF@CCUBE A,                                                                                                                                                                                                               \n";
            strSQL = strSQL + "                        (                                                                                                                                                                                                                                            \n";
            strSQL = strSQL + "                         SELECT SAWON_NO,MAX(SEQ)SEQ                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "                         FROM FTUSER.SAWON_PAY_INF@CCUBE                                                                                                                                                                                                             \n";
            strSQL = strSQL + "                         WHERE DEL_FLAG IS NULL                                                                                                                                                                                                                      \n";
            strSQL = strSQL + "                         AND (E_DATE IS NULL OR E_DATE > '" + strCrntDT + "') AND S_DATE <= '" + strCrntDT + "'                                                                                                                                                      \n";
            strSQL = strSQL + "                         GROUP BY SAWON_NO                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                        )B                                                                                                                                                                                                                                           \n";
            strSQL = strSQL + "                    WHERE A.SAWON_NO = B.SAWON_NO                                                                                                                                                                                                                    \n";
            strSQL = strSQL + "                    AND A.SEQ = B.SEQ                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "                   )D                                                                                                                                                                                                                                                \n";
            //인상시급 관련 쿼리 적용 (조회날짜에 따라 인상된 시급이 변경되어 조회되게끔 수정)                                                                                                                                                                                                                                     
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_GROUP,MANAGE2 AS S_GROUP_SEQ FROM PM000STD@CCUBE WHERE MAIN_CODE = '00067')E                                                                                                                                           \n";
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_JO FROM PM000STD@CCUBE WHERE MAIN_CODE ='00064')F                                                                                                                                                                      \n";
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_BAY,MANAGE5 AS S_BAY_SEQ FROM PM000STD@CCUBE WHERE MAIN_CODE ='00066')G                                                                                                                                                \n";
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_OS FROM PM000STD@CCUBE WHERE MAIN_CODE = '00096')H                                                                                                                                                                     \n";
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_COST FROM PM000STD@CCUBE WHERE MAIN_CODE = '00099')I                                                                                                                                                                   \n";
            strSQL = strSQL + "              ,(SELECT SUB_CODE,SUB_NAME AS S_DEPT FROM PM000STD@CCUBE WHERE MAIN_CODE = '00118')J                                                                                                                                                                   \n";
            strSQL = strSQL + "    WHERE A.SAWON_NO = B.SAWON_NO                                                                                                                                                                                                                                    \n";
            //strSQL = strSQL + "    AND A.SAWON_NO = C.SAWON_NO(+)   --//트랜드 쿼리는 근태테이블과 인사정보테이블을 INNER JOIN으로 걸었음, 기존 협의 시 근태등록한 데이터를 기준으로 한다고 하여 OUTER JOIN을 INNER JOIN으로 변경                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "    AND A.SAWON_NO = C.SAWON_NO                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "    AND A.SAWON_NO = D.SAWON_NO(+)                                                                                                                                                                                                                                   \n";
            strSQL = strSQL + "    AND B.ASSIGN_GROUP = E.SUB_CODE(+)                                                                                                                                                                                                                               \n";
            strSQL = strSQL + "    AND B.ASSIGN_JOIRUM = F.SUB_CODE(+)                                                                                                                                                                                                                              \n";
            strSQL = strSQL + "    AND B.ASSIGN_BAY = G.SUB_CODE(+)                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "    AND B.SAWON_OUTSOURCING = H.SUB_CODE(+)                                                                                                                                                                                                                          \n";
            strSQL = strSQL + "    AND B.COST_GROUP =  I.SUB_CODE(+)                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "    AND B.ASSIGN_DEPT = J.SUB_CODE(+)                                                                                                                                                                                                                                \n";
            strSQL = strSQL + "   )                                                                                                                                                                                                                                                                 \n";
            strSQL = strSQL + "GROUP BY ASSIGN_GROUP,ASSIGN_PART,ASSIGN_BAY,ASSIGN_JOIRUM,SAWON_OUTSOURCING ,S_GROUP,S_GROUP_SEQ,S_JO,S_BAY,S_BAY_SEQ,S_OS,S_COST,S_DEPT)                                                                                                                           \n";
            strSQL = strSQL + " WHERE 1=1                                                                                                                                                                                                                                                           \n";

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                strSQL = strSQL + "AND DEPT ='" + dr_dept.SelectedValue + "' \n";
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Area.Count > 0)
                if (mcc_dr_Area.Count > 0 && mcc_dr_Area.Text != "")
                {
                    strSQL = strSQL + "  AND AREA IN (" + mcc_dr_Area.SQLText + ") \n";
                }
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Part.Count > 0)
                if (mcc_dr_Part.Count > 0 && mcc_dr_Part.Text != "")
                {
                    strSQL = strSQL + "  AND PART IN (" + mcc_dr_Part.SQLText + ") \n";
                }
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Oprgrp.Count > 0)
                if (mcc_dr_Oprgrp.Count > 0 && mcc_dr_Oprgrp.Text != "")
                {
                    strSQL = strSQL + "  AND OPER_GROUP IN (" + mcc_dr_Oprgrp.SQLText + ") \n";
                }
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Cost.Count > 0)
                if (mcc_dr_Cost.Count > 0 && mcc_dr_Cost.Text != "")
                {
                    strSQL = strSQL + "  AND S_COST IN (" + mcc_dr_Cost.SQLText + ") \n";
                }
            }

            if (dr_dept.SelectedValue == "Display" || dr_dept.SelectedValue == "Semi")
            {
                //if (mcc_dr_Depart.Count > 0)
                if (mcc_dr_Depart.Count > 0 && mcc_dr_Depart.Text != "")
                {
                    strSQL = strSQL + "  AND S_DEPT IN (" + mcc_dr_Depart.SQLText + ") \n";
                }
            }

            strSQL = strSQL + " ORDER BY S_DEPT, AREA , PART, S_COST, OPER_GROUP \n";

            return strSQL;
        }
        #endregion

        #region "Event"
        protected void InitTxtWorkDate(object sender, EventArgs e)
        {
            TextBox txtTemp = (TextBox)sender;
            txtTemp.Text = (DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
        }

        //protected void Page_PreInit(Object sender, EventArgs e)
        //{
        //    if (Request.QueryString["Menu"] == "false") this.MasterPageFile = "../Site2.Master";
        //}

        protected void Page_Load(object sender, EventArgs e)
        {

            SetDefaultValue();

            if (!IsPostBack)
            {
                Reprt_Reset(sender, e);
                WebSiteCount();
            }


            if (Request.QueryString["Menu"] == "false")
            {
                //SiteMapPath1.Visible = false;
                Table1.Visible = true;
            }
        }

        public void WebSiteCount()
        {
            cls_useSiteCount siteCount = new cls_useSiteCount();
            siteCount.strDB = "NEPES&NEPES_DISPLAY";
            siteCount.strMenu = HttpContext.Current.Request.Url.AbsolutePath.ToString();
            siteCount.SetMenuCountInsert();
        }

        protected void query_Click(object sender, EventArgs e)
        {
            string strCrntDT = txtCrntDT.Text, strdeptID = "", strSQL = "";

            if (txtCrntDT.Text.Trim().Length == 0) return;
            if ((strSQL = MakeQuery(strCrntDT, strdeptID)) == "") return;

            ERPAppAddition.ERPAddition.INSA.DailyPaySum.CommuteModule dtHM1 = new ERPAppAddition.ERPAddition.INSA.DailyPaySum.CommuteModule();
            GF.CreateReport(dtHM1, 0, strSQL, ReportViewer1, "INSA.DailyPaySum.DailyPaySum.rdlc", "dsDailyPaySum", "REPORT_" + this.Title + DateTime.Now.ToShortDateString());
        }

        protected void Reprt_Reset(object sender, EventArgs e)
        {
            ReportViewer1.Reset();

            mcc_dr_Area.ClearSQLText();
            mcc_dr_Part.ClearSQLText();
            mcc_dr_Oprgrp.ClearSQLText();
            mcc_dr_Cost.ClearSQLText(); // 151114 변수초기화 추가
            mcc_dr_Depart.ClearSQLText();  // 151114 변수초기화 추가
        }


        #endregion

    }
}