using System;
using System.Runtime.InteropServices;

namespace ERPAppAddition.ERPAddition
{
    public class GC
    {
        public const string reportPath = "ERPAppAddition.ERPAddition.";
        public enum ConnType : byte { DISP, SEMI };
        public enum DateTimeFmt : byte { yyyy, yyyyMM, yyyyMMdd, yyyyMMddhh, yyyyMMddhhmm, yyyyMMddhhmmss };

        #region "SQL_FOR_DISPLAY"
        // for Display
        public enum MatType : byte { COMP, MTRL, SHET };
        public const string SQL_MAT_TYPE_COMP = "('CM', 'RC')";
        public const string SQL_MAT_ID = "SELECT MAT_ID FROM MWIPMATDEF WHERE FACTORY = 'DISPLAY'";
        public const string SQL_MAT_ID_COMP = SQL_MAT_ID + " AND MAT_TYPE IN ('CM','RC') AND DELETE_FLAG = ' ' ORDER BY MAT_ID";
        public const string SQL_MAT_ID_MTRL = SQL_MAT_ID + " AND MAT_TYPE IN ('RM') AND DELETE_FLAG = ' ' ORDER BY MAT_ID";
        public const string SQL_MAT_ID_SHET = SQL_MAT_ID + " AND UNIT_1 = 'SHEET' AND DELETE_FLAG = ' ' ORDER BY MAT_ID";
        public const string SQL_SFLOW = "SELECT FLOW, FLOW||' '||FLOW_DESC AS FLOW_DESC FROM MWIPFLWDEF WHERE FACTORY = 'DISPLAY' ORDER BY 1";
        public const string SQL_OPER_COMP = "SELECT OPER||' '||OPER_DESC AS OPER FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' ORDER BY 1";
        public const string SQL_OPER_MTRL = "SELECT OPER||' '||OPER_DESC AS OPER FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' AND INV_FLAG = 'Y' ORDER BY 1";
        public const string SQL_SOPER_COMP = "SELECT OPER, OPER||' '||OPER_DESC AS OPER_DESC FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' ORDER BY 1";
        public const string SQL_SOPER_MTRL = "SELECT OPER, OPER||' '||OPER_DESC AS OPER_DESC FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' AND INV_FLAG = 'Y' ORDER BY 1";
        public const string SQL_LOSS_ALL = "SELECT KEY_1||' '||DATA_1 AS LOSSCD FROM MGCMTBLDAT WHERE FACTORY = 'DISPLAY' AND TABLE_NAME ='DEFECT_CODE' ORDER BY 1";
        public const string SQL_SLOSS_ALL = "SELECT KEY_1 AS LOSS_CODE, KEY_1||' '||DATA_1 AS LOSS_DESC FROM MGCMTBLDAT WHERE FACTORY = 'DISPLAY' AND TABLE_NAME ='DEFECT_CODE' ORDER BY 1";
        public const string SQL_SCREATE_CODE = "SELECT KEY_1 CREATE_CODE, KEY_1||' '||DATA_1 AS CC_DESC FROM MGCMTBLDAT WHERE FACTORY = 'DISPLAY' AND TABLE_NAME = 'CREATE_CODE'";
        public const string SQL_ENTR_OPER = "SELECT OPER FROM MWIPOPRDEF WHERE FACTORY = 'DISPLAY' AND OPER_GRP_1 = '입고'";
        #endregion

        #region "SQL_FOR_SEMICONDUCTOR"
        public enum PlantType : byte { ALL = 0, P01, P02, P09, P12 };
        public enum ProdType : byte { ALL = 0, DDI, WLP, FOWLP };
        public enum MethodType : byte { CRNT, STND };
        public const string SQL_PART_ID = "SELECT DISTINCT PART AS PART_ID FROM REPORTWIP";
        public const string SQL_SOPER = "SELECT DISTINCT OPERATION AS OPER, OPERATION||' '||(SELECT SHORT_DESC FROM OPERATION WHERE OPERATION = RW.OPERATION AND PLANT = 'CCUBEDIGITAL') AS OPER_DESC FROM REPORTWIP RW";
        public const string SQL_CREATE_CODE = "SELECT DISTINCT CREATE_CODE AS CREATE_CODE FROM REPORTWIP";
        //public const string BSQL_CRNT = "FROM REPORTWIP WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 ORDER BY 1";

        public const string SQL_CRNT_P01 = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND PROD_TYPE IN ('6', '8') AND QTY_UNIT_1 = 'SLS' ORDER BY 1";
        public const string SQL_CRNT_P02 = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND QTY_UNIT_1 = 'PCS' ORDER BY 1";
        public const string SQL_CRNT_P09 = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND PROD_TYPE = 'C' AND QTY_UNIT_1 = 'SLS' ORDER BY 1";
        public const string SQL_CRNT_P12 = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND CREATE_CODE LIKE 'RCP%' ORDER BY 1";

        public const string SQL_CRNT_DDI = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND CREATE_CODE NOT LIKE 'RCP%' AND CREATE_CODE NOT IN ('WSLD', 'WLCSP') ORDER BY 1";
        public const string SQL_CRNT_WLP = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND CREATE_CODE IN ('WSLD', 'WLCSP') ORDER BY 1";
        public const string SQL_CRNT_FOWLP = " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 AND CREATE_CODE LIKE 'RCP%' ORDER BY 1";

        public const string SQL_PART_ID_CRNT = SQL_PART_ID + " WHERE PLANT = 'CCUBEDIGITAL' AND STATUS <> 99 ORDER BY 1";
        public const string SQL_PART_ID_STND = "SELECT DISTINCT PART_ID AS PART_ID FROM MIGHTY.PART@CCUBE WHERE PLANT = 'CCUBEDIGITAL' AND APPROVAL = 'Y' ORDER BY 1";

        public const string SQL_PART_ID_CRNT_P01 = SQL_PART_ID + SQL_CRNT_P01;
        public const string SQL_PART_ID_CRNT_P02 = SQL_PART_ID + SQL_CRNT_P02;
        public const string SQL_PART_ID_CRNT_P09 = SQL_PART_ID + SQL_CRNT_P09;
        public const string SQL_PART_ID_CRNT_P12 = SQL_PART_ID + SQL_CRNT_P12;
        public const string SQL_SOPER_CRNT_P01 = SQL_SOPER + SQL_CRNT_P01;
        public const string SQL_SOPER_CRNT_P02 = SQL_SOPER + SQL_CRNT_P02;
        public const string SQL_SOPER_CRNT_P09 = SQL_SOPER + SQL_CRNT_P09;
        public const string SQL_SOPER_CRNT_P12 = SQL_SOPER + SQL_CRNT_P12;
        public const string SQL_CREATE_CODE_CRNT_P01 = SQL_CREATE_CODE + SQL_CRNT_P01;
        public const string SQL_CREATE_CODE_CRNT_P02 = SQL_CREATE_CODE + SQL_CRNT_P02;
        public const string SQL_CREATE_CODE_CRNT_P09 = SQL_CREATE_CODE + SQL_CRNT_P09;
        public const string SQL_CREATE_CODE_CRNT_P12 = SQL_CREATE_CODE + SQL_CRNT_P12;

        public const string SQL_PART_ID_CRNT_DDI = SQL_PART_ID + SQL_CRNT_DDI;
        public const string SQL_PART_ID_CRNT_WLP = SQL_PART_ID + SQL_CRNT_WLP;
        public const string SQL_PART_ID_CRNT_FOWLP = SQL_PART_ID + SQL_CRNT_FOWLP;
        public const string SQL_SOPER_CRNT_DDI = SQL_SOPER + SQL_CRNT_DDI;
        public const string SQL_SOPER_CRNT_WLP = SQL_SOPER + SQL_CRNT_WLP;
        public const string SQL_SOPER_CRNT_FOWLP = SQL_SOPER + SQL_CRNT_FOWLP;
        public const string SQL_CREATE_CODE_CRNT_DDI = SQL_CREATE_CODE + SQL_CRNT_DDI;
        public const string SQL_CREATE_CODE_CRNT_WLP = SQL_CREATE_CODE + SQL_CRNT_WLP;
        public const string SQL_CREATE_CODE_CRNT_FOWLP = SQL_CREATE_CODE + SQL_CRNT_FOWLP;
        #endregion

        #region "SQL_FOR_HRMS"
        // 인건비레포트 조회조건 관련 추가 (2015.06.17_안창주)
        public const string SQL_AREA_CODE_SEMI = @"SELECT DISTINCT
                                                     B.SUB_NAME AS AREA
                                                   , B.SUB_CODE AS AREA_CODE
                                                   FROM SAWON_WORK_TIME_INF@CCUBE A, PM000STD@CCUBE B
                                                     WHERE 1=1
                                                      AND B.SUB_CODE = A.SAWON_GROUP
                                                      AND B.MAIN_CODE = '00067' 
                                                      AND B.SUB_NAME NOT LIKE 'HRMS%'
                                                      AND B.SUB_CODE <> '00000'";
        public const string SQL_OPERGRP_CODE_SEMI = @"SELECT DISTINCT
                                                        B.SUB_NAME AS OPER_GROUP
                                                      , B.SUB_CODE AS OPER_CODE    
                                                      FROM SAWON_WORK_TIME_INF@CCUBE A,  PM000STD@CCUBE B
                                                        WHERE 1=1
                                                         AND B.SUB_CODE = A.SAWON_BAY
                                                         AND B.MAIN_CODE = '00066' 
                                                         AND B.SUB_NAME NOT LIKE 'HRMS%'
                                                         AND B.SUB_CODE <> '00000'";
        public const string SQL_PART_CODE_SEMI = @"SELECT DISTINCT
                                                     MANAGE1 AS PART
                                                   FROM PM000STD@CCUBE
                                                     WHERE 1=1
                                                      AND MAIN_CODE = '00066'
                                                      AND MANAGE2 IS NOT NULL
                                                      AND SUB_CODE <> '000000'
                                                      AND MANAGE1 <> 'Part 구분'";
        public const string SQL_COST_CODE_SEMI = @"SELECT SUB_CODE
                                                        , SUB_NAME AS S_GROUP 
                                                   FROM PM000STD@CCUBE 
                                                     WHERE 1=1
                                                      AND MAIN_CODE = '00099' 
                                                      AND SUB_CODE <> '00000'"; // 150816 신규추가
        public const string SQL_DEPART_CODE_SEMI = @"SELECT SUB_CODE
                                                          , SUB_NAME AS S_GROUP 
                                                     FROM PM000STD@CCUBE 
                                                       WHERE 1=1
                                                        AND MAIN_CODE = '00118' 
                                                        AND SUB_CODE <> '00000'"; // 150816 신규추가
        public const string SQL_AREA_CODE_DISP = "SELECT DISTINCT AREA FROM NHRMWRKIOT";
        public const string SQL_OPERGRP_CODE_DISP = "SELECT DISTINCT OPER_GROUP FROM NHRMWRKIOT";
        #endregion

        #region "Transaction Code"
        public const string TRAN_CODE_CREATE = "CREATE";
        public const string TRAN_CODE_START = "START";
        public const string TRAN_CODE_END = "END";
        public const string TRAN_CODE_MOVE = "MOVE";
        public const string TRAN_CODE_SKIP = "SKIP";
        public const string TRAN_CODE_REWORK = "REWORK";
        public const string TRAN_CODE_REPAIR = "REPAIR";
        public const string TRAN_CODE_REPAIR_END = "REPAIR_END";
        public const string TRAN_CODE_LOCAL_REPAIR = "LOCAL_REPAIR";
        public const string TRAN_CODE_LOSS = "LOSS";
        public const string TRAN_CODE_BONUS = "BONUS";
        public const string TRAN_CODE_SPLIT = "SPLIT";
        public const string TRAN_CODE_MERGE = "MERGE";
        public const string TRAN_CODE_COMBINE = "COMBINE";
        public const string TRAN_CODE_HOLD = "HOLD";
        public const string TRAN_CODE_RELEASE = "RELEASE";
        public const string TRAN_CODE_SHIP = "SHIP";
        public const string TRAN_CODE_RECEIVE = "RECEIVE";
        public const string TRAN_CODE_ADAPT = "ADAPT";
        public const string TRAN_CODE_ATTRIBUTE = "ATTRIBUTE";
        public const string TRAN_CODE_LOTEDC = "LOTEDC";
        public const string TRAN_CODE_RESEDC = "RESEDC";
        public const string TRAN_CODE_SORT = "SORT";

        public const string TRAN_CODE_IN_INV = "IN INV";
        public const string TRAN_CODE_OUT_INV = "OUT INV";
        public const string TRAN_CODE_TRANSFER_INV = "TRANS INV";
        public const string TRAN_CODE_CONV_TO_LOT = "CONV TO LOT";
        public const string TRAN_CODE_CONV_TO_INV = "CONV TO INV";
        public const string TRAN_CODE_CONSUME = "CONSUME";
        public const string TRAN_CODE_SCRAP = "SCRAP";
        public const string TRAN_CODE_REVERSE = "REVERSE";
        public const string TRAN_CODE_SCRIBE = "SCRIBE";

        public const string TRAN_CODE_TERMINATE = "TERMINATE";
        public const string TRAN_CODE_RECREATE = "RECREATE";
        public const string TRAN_CODE_CLEAN = "CLEAN";
        #endregion

        #region "Process Step"
        public const char STEP_CREATE = 'I';
        public const char STEP_UPDATE = 'U';
        public const char STEP_DELETE = 'D';
        public const char STEP_CONFIRM = 'F';
        public const char STEP_DELETE_FORCE = 'X';
        public const char STEP_COPY = 'C';
        public const char STEP_UNDELETE = 'R';
        public const char STEP_APPROVAL = 'A';
        public const char STEP_RELEASE = 'E';
        public const char STEP_CANCEL_APPROVAL = 'P';
        public const char STEP_SCRAP = 'S';
        public const char STEP_RETURN = 'N';
        public const char STEP_REGENERATE = 'G';
        public const char STEP_VERSION_UP = 'V';
        public const char STEP_TERMINATE = 'M';
        public const char STEP_TRAN = 'T';
        #endregion

        #region "CMF Item"
        //CMF
        public const string CMF_MATERIAL = "CMF_MATERIAL";
        public const string CMF_FLOW = "CMF_FLOW";
        public const string CMF_OPERATION = "CMF_OPER";
        public const string CMF_STEP = "CMF_STEP";
        public const string CMF_RESOURCE = "CMF_RESOURCE";
        public const string CMF_PORT = "CMF_PORT";

        public const string CMF_CARRIER = "CMF_CARRIER";
        public const string CMF_SUBRESOURCE = "CMF_SUBRESOURCE";
        public const string CMF_USER = "CMF_USER";
        public const string CMF_BOM_PART = "CMF_BOM_PART";
        public const string CMF_BOM_ORDER = "CMF_BOM_ORDER";
        public const string CMF_CHARACTER = "CMF_CHARACTER";
        public const string CMF_COL_SET = "CMF_COL_SET";
        public const string CMF_DFO = "CMF_MFO";
        public const string CMF_CALENDAR = "CMF_CALENDAR";
        public const string CMF_EVENT = "CMF_EVENT";
        public const string CMF_ORDER = "CMF_ORDER";
        public const string CMF_BOM_SET = "CMF_BOM_SET";
        public const string CMF_RECIPE = "CMF_RECIPE";
        public const string CMF_PART = "CMF_PART";
        public const string CMF_LABEL = "CMF_LABEL";
        public const string CMF_LOT_ALARM = "CMF_LOT_ALARM";
        public const string CMF_RES_ALARM = "CMF_RES_ALARM";

        public const string CMF_QUEUETIME = "CMF_QUEUETIME";
        public const string CMF_SERVICE = "CMF_SERVICE";

        public const string CMF_LOT = "CMF_LOT";
        public const string CMF_SUBLOT = "CMF_SUBLOT";

        public const string CMF_TRN_ADAPT = "CMF_TRN_ADAPT";
        public const string CMF_TRN_ATTRIBUTE = "CMF_TRN_ATTRIBUTE";
        public const string CMF_TRN_BONUS = "CMF_TRN_BONUS";
        public const string CMF_TRN_LOSS = "CMF_TRN_LOSS";
        public const string CMF_TRN_CREATE = "CMF_TRN_CREATE";
        public const string CMF_TRN_START = "CMF_TRN_START";
        public const string CMF_TRN_END = "CMF_TRN_END";
        public const string CMF_TRN_MOVE = "CMF_TRN_MOVE";
        public const string CMF_TRN_SKIP = "CMF_TRN_SKIP";
        public const string CMF_TRN_REWORK = "CMF_TRN_REWORK";
        public const string CMF_TRN_REPAIR = "CMF_TRN_REPAIR";
        public const string CMF_TRN_REPAIR_END = "CMF_TRN_REPAIR_END";
        public const string CMF_TRN_LOCAL_REPAIR = "CMF_TRN_LOCAL_REPAIR";
        public const string CMF_TRN_SPLIT = "CMF_TRN_SPLIT";
        public const string CMF_TRN_COMBINE = "CMF_TRN_COMBINE";
        public const string CMF_TRN_MERGE = "CMF_TRN_MERGE";
        public const string CMF_TRN_HOLD = "CMF_TRN_HOLD";
        public const string CMF_TRN_RELEASE = "CMF_TRN_RELEASE";
        public const string CMF_TRN_SHIP = "CMF_TRN_SHIP";
        public const string CMF_TRN_RECEIVE = "CMF_TRN_RECEIVE";
        public const string CMF_TRN_ASSEMBLY = "CMF_TRN_ASSEMBLY";
        public const string CMF_TRN_DISASSEMBLE = "CMF_TRN_DISASSEMBLE";
        public const string CMF_TRN_REPLACE = "CMF_TRN_REPLACE";
        public const string CMF_TRN_LOTEDC = "CMF_TRN_LOTEDC";
        public const string CMF_TRN_EVENT = "CMF_TRN_EVENT";
        public const string CMF_TRN_TROUBLE = "CMF_TRN_TROUBLE";
        public const string CMF_TRN_RMA_OPEN = "CMF_TRN_RMA_OPEN";
        public const string CMF_TRN_RMA_CLOSE = "CMF_TRN_RMA_CLOSE";
        public const string CMF_TRN_SORT = "CMF_TRN_SORT";
        public const string CMF_TRN_STORE = "CMF_TRN_STORE";
        public const string CMF_TRN_UNSTORE = "CMF_TRN_UNSTORE";
        public const string CMF_TRN_TERMINATE = "CMF_TRN_TERMINATE";
        public const string CMF_TRN_CHANGE_CMF = "CMF_TRN_CHANGE_CMF";
        public const string CMF_TRN_RESERVE = "CMF_TRN_RESERVE";
        public const string CMF_TRN_UNRESERVE = "CMF_TRN_UNRESERVE";
        public const string CMF_TRN_GRADE = "CMF_TRN_GRADE";
        public const string CMF_TRN_SCRIBE = "CMF_TRN_SCRIBE";
        public const string CMF_TRN_CV = "CMF_TRN_CV";
        public const string CMF_TRN_REGENERATE = "CMF_TRN_REGENERATE";

        //change port status
        public const string CMF_TRN_CHANGE_PORT = "CMF_TRN_CHANGE_PORT";
        public const string CMF_TRN_COLLECT_LOT_DEFECT = "CMF_TRN_COLLECT_DFT";
        public const string CMF_TRN_CLEAN_LOT_DEFECT = "CMF_TRN_CLEAN_DFT";

        //Inventory CMF
        public const string CMF_TRN_IN_INV = "CMF_TRN_IN_INV";
        public const string CMF_TRN_OUT_INV = "CMF_TRN_OUT_INV";
        public const string CMF_TRN_TRANS_INV = "CMF_TRN_TRANS_INV";
        public const string CMF_TRN_CONV_TO_LOT = "CMF_TRN_CONV_TO_LOT";
        public const string CMF_TRN_CONV_TO_INV = "CMF_TRN_CONV_TO_INV";
        public const string CMF_TRN_CONSUME = "CMF_TRN_CONSUME";
        public const string CMF_TRN_SCRAP = "CMF_TRN_SCRAP";

        public const string CMF_TRN_QCM_BATCH = "CMF_TRN_QCM_BATCH";
        public const string CMF_TRN_QCM_RESULT = "CMF_TRN_QCM_RESULT";
        public const string CMF_TRN_QCM_FINAL = "CMF_TRN_QCM_FINAL";
        public const string CMF_TRN_QCM_MERGE = "CMF_TRN_QCM_MERGE";
        public const string CMF_TRN_QCM_SPLIT = "CMF_TRN_QCM_SPLIT";
        public const string CMF_CHART_SET = "CMF_CHART_SET";
        public const string CMF_CHART = "CMF_CHART";

        public const string CMF_RULE_RELATION = "CMF_RULE_RELATION";
        public const string CMF_RULE_SEQ_KEY = "CMF_RULE_SEQ_KEY";

        //Group
        public const string GRP_FLOW = "GRP_FLOW";
        public const string GRP_MATERIAL = "GRP_MATERIAL";
        public const string GRP_OPERATION = "GRP_OPER";
        public const string GRP_STEP = "GRP_STEP";
        public const string GRP_CHARACTER = "GRP_CHARACTER";
        public const string GRP_RESOURCE = "GRP_RESOURCE";
        public const string GRP_COL_SET = "GRP_COL_SET";
        public const string GRP_USER = "GRP_USER";
        public const string GRP_EVENT = "GRP_EVENT";
        public const string GRP_BOM_SET = "GRP_BOM_SET";
        public const string GRP_RECIPE = "GRP_RECIPE";
        public const string GRP_INSP_SET = "GRP_INSP_SET";
        public const string GRP_CHART = "GRP_CHART";
        public const string GRP_CHART_SET = "GRP_CHART_SET";
        #endregion

        #region "GCM Table Name"
        //System GCM Table
        public const string GCM_MSGGRP_TBL = "MESSAGE_GROUP";

        //Collection Set Group Table 1~10
        public const string GCM_COL_GRP_1 = "COL_GRP_1";
        public const string GCM_COL_GRP_2 = "COL_GRP_2";
        public const string GCM_COL_GRP_3 = "COL_GRP_3";
        public const string GCM_COL_GRP_4 = "COL_GRP_4";
        public const string GCM_COL_GRP_5 = "COL_GRP_5";
        public const string GCM_COL_GRP_6 = "COL_GRP_6";
        public const string GCM_COL_GRP_7 = "COL_GRP_7";
        public const string GCM_COL_GRP_8 = "COL_GRP_8";
        public const string GCM_COL_GRP_9 = "COL_GRP_9";
        public const string GCM_COL_GRP_10 = "COL_GRP_10";

        //Character Group Table 1~10
        public const string GCM_CHAR_GRP_1 = "CHAR_GRP_1";
        public const string GCM_CHAR_GRP_2 = "CHAR_GRP_2";
        public const string GCM_CHAR_GRP_3 = "CHAR_GRP_3";
        public const string GCM_CHAR_GRP_4 = "CHAR_GRP_4";
        public const string GCM_CHAR_GRP_5 = "CHAR_GRP_5";
        public const string GCM_CHAR_GRP_6 = "CHAR_GRP_6";
        public const string GCM_CHAR_GRP_7 = "CHAR_GRP_7";
        public const string GCM_CHAR_GRP_8 = "CHAR_GRP_8";
        public const string GCM_CHAR_GRP_9 = "CHAR_GRP_9";
        public const string GCM_CHAR_GRP_10 = "CHAR_GRP_10";

        //Resource Group Table 1~10
        public const string GCM_RES_GRP_1 = "RES_GRP_1";
        public const string GCM_RES_GRP_2 = "RES_GRP_2";
        public const string GCM_RES_GRP_3 = "RES_GRP_3";
        public const string GCM_RES_GRP_4 = "RES_GRP_4";
        public const string GCM_RES_GRP_5 = "RES_GRP_5";
        public const string GCM_RES_GRP_6 = "RES_GRP_6";
        public const string GCM_RES_GRP_7 = "RES_GRP_7";
        public const string GCM_RES_GRP_8 = "RES_GRP_8";
        public const string GCM_RES_GRP_9 = "RES_GRP_9";
        public const string GCM_RES_GRP_10 = "RES_GRP_10";

        //Event Group Table 1~10
        public const string GCM_EVN_GRP_1 = "EVN_GRP_1";
        public const string GCM_EVN_GRP_2 = "EVN_GRP_2";
        public const string GCM_EVN_GRP_3 = "EVN_GRP_3";
        public const string GCM_EVN_GRP_4 = "EVN_GRP_4";
        public const string GCM_EVN_GRP_5 = "EVN_GRP_5";
        public const string GCM_EVN_GRP_6 = "EVN_GRP_6";
        public const string GCM_EVN_GRP_7 = "EVN_GRP_7";
        public const string GCM_EVN_GRP_8 = "EVN_GRP_8";
        public const string GCM_EVN_GRP_9 = "EVN_GRP_9";
        public const string GCM_EVN_GRP_10 = "EVN_GRP_10";

        //Material Group Table 1~10
        public const string GCM_MATERIAL_GRP_1 = "MATERIAL_GRP_1";
        public const string GCM_MATERIAL_GRP_2 = "MATERIAL_GRP_2";
        public const string GCM_MATERIAL_GRP_3 = "MATERIAL_GRP_3";
        public const string GCM_MATERIAL_GRP_4 = "MATERIAL_GRP_4";
        public const string GCM_MATERIAL_GRP_5 = "MATERIAL_GRP_5";
        public const string GCM_MATERIAL_GRP_6 = "MATERIAL_GRP_6";
        public const string GCM_MATERIAL_GRP_7 = "MATERIAL_GRP_7";
        public const string GCM_MATERIAL_GRP_8 = "MATERIAL_GRP_8";
        public const string GCM_MATERIAL_GRP_9 = "MATERIAL_GRP_9";
        public const string GCM_MATERIAL_GRP_10 = "MATERIAL_GRP_10";

        //Flow Group Table 1~10
        public const string GCM_FLOW_GRP_1 = "FLOW_GRP_1";
        public const string GCM_FLOW_GRP_2 = "FLOW_GRP_2";
        public const string GCM_FLOW_GRP_3 = "FLOW_GRP_3";
        public const string GCM_FLOW_GRP_4 = "FLOW_GRP_4";
        public const string GCM_FLOW_GRP_5 = "FLOW_GRP_5";
        public const string GCM_FLOW_GRP_6 = "FLOW_GRP_6";
        public const string GCM_FLOW_GRP_7 = "FLOW_GRP_7";
        public const string GCM_FLOW_GRP_8 = "FLOW_GRP_8";
        public const string GCM_FLOW_GRP_9 = "FLOW_GRP_9";
        public const string GCM_FLOW_GRP_10 = "FLOW_GRP_10";

        //Operation Group Table 1~10
        public const string GCM_OPER_GRP_1 = "OPER_GRP_1";
        public const string GCM_OPER_GRP_2 = "OPER_GRP_2";
        public const string GCM_OPER_GRP_3 = "OPER_GRP_3";
        public const string GCM_OPER_GRP_4 = "OPER_GRP_4";
        public const string GCM_OPER_GRP_5 = "OPER_GRP_5";
        public const string GCM_OPER_GRP_6 = "OPER_GRP_6";
        public const string GCM_OPER_GRP_7 = "OPER_GRP_7";
        public const string GCM_OPER_GRP_8 = "OPER_GRP_8";
        public const string GCM_OPER_GRP_9 = "OPER_GRP_9";
        public const string GCM_OPER_GRP_10 = "OPER_GRP_10";

        //Step Group Table 1~10
        public const string GCM_STEP_GRP_1 = "STEP_GRP_1";
        public const string GCM_STEP_GRP_2 = "STEP_GRP_2";
        public const string GCM_STEP_GRP_3 = "STEP_GRP_3";
        public const string GCM_STEP_GRP_4 = "STEP_GRP_4";
        public const string GCM_STEP_GRP_5 = "STEP_GRP_5";
        public const string GCM_STEP_GRP_6 = "STEP_GRP_6";
        public const string GCM_STEP_GRP_7 = "STEP_GRP_7";
        public const string GCM_STEP_GRP_8 = "STEP_GRP_8";
        public const string GCM_STEP_GRP_9 = "STEP_GRP_9";
        public const string GCM_STEP_GRP_10 = "STEP_GRP_10";

        //User Group Table 1~10
        public const string GCM_USER_GRP_1 = "USER_GRP_1";
        public const string GCM_USER_GRP_2 = "USER_GRP_2";
        public const string GCM_USER_GRP_3 = "USER_GRP_3";
        public const string GCM_USER_GRP_4 = "USER_GRP_4";
        public const string GCM_USER_GRP_5 = "USER_GRP_5";
        public const string GCM_USER_GRP_6 = "USER_GRP_6";
        public const string GCM_USER_GRP_7 = "USER_GRP_7";
        public const string GCM_USER_GRP_8 = "USER_GRP_8";
        public const string GCM_USER_GRP_9 = "USER_GRP_9";
        public const string GCM_USER_GRP_10 = "USER_GRP_10";


        //BOM Set Group Table 1~10
        public const string GCM_BOM_GRP_1 = "BOM_GRP_1";
        public const string GCM_BOM_GRP_2 = "BOM_GRP_2";
        public const string GCM_BOM_GRP_3 = "BOM_GRP_3";
        public const string GCM_BOM_GRP_4 = "BOM_GRP_4";
        public const string GCM_BOM_GRP_5 = "BOM_GRP_5";
        public const string GCM_BOM_GRP_6 = "BOM_GRP_6";
        public const string GCM_BOM_GRP_7 = "BOM_GRP_7";
        public const string GCM_BOM_GRP_8 = "BOM_GRP_8";
        public const string GCM_BOM_GRP_9 = "BOM_GRP_9";
        public const string GCM_BOM_GRP_10 = "BOM_GRP_10";

        //Recipe Group Table 1~10
        public const string GCM_RECIPE_GRP_1 = "RECIPE_GRP_1";
        public const string GCM_RECIPE_GRP_2 = "RECIPE_GRP_2";
        public const string GCM_RECIPE_GRP_3 = "RECIPE_GRP_3";
        public const string GCM_RECIPE_GRP_4 = "RECIPE_GRP_4";
        public const string GCM_RECIPE_GRP_5 = "RECIPE_GRP_5";
        public const string GCM_RECIPE_GRP_6 = "RECIPE_GRP_6";
        public const string GCM_RECIPE_GRP_7 = "RECIPE_GRP_7";
        public const string GCM_RECIPE_GRP_8 = "RECIPE_GRP_8";
        public const string GCM_RECIPE_GRP_9 = "RECIPE_GRP_9";
        public const string GCM_RECIPE_GRP_10 = "RECIPE_GRP_10";

        //Inspection Set Group Table 1~10
        public const string GCM_INSP_SET_GRP_1 = "INSP_SET_GRP_1";
        public const string GCM_INSP_SET_GRP_2 = "INSP_SET_GRP_2";
        public const string GCM_INSP_SET_GRP_3 = "INSP_SET_GRP_3";
        public const string GCM_INSP_SET_GRP_4 = "INSP_SET_GRP_4";
        public const string GCM_INSP_SET_GRP_5 = "INSP_SET_GRP_5";
        public const string GCM_INSP_SET_GRP_6 = "INSP_SET_GRP_6";
        public const string GCM_INSP_SET_GRP_7 = "INSP_SET_GRP_7";
        public const string GCM_INSP_SET_GRP_8 = "INSP_SET_GRP_8";
        public const string GCM_INSP_SET_GRP_9 = "INSP_SET_GRP_9";
        public const string GCM_INSP_SET_GRP_10 = "INSP_SET_GRP_10";

        //SPC Chart Group Table 1~10
        public const string GCM_CHT_GRP_1 = "CHT_GRP_1";
        public const string GCM_CHT_GRP_2 = "CHT_GRP_2";
        public const string GCM_CHT_GRP_3 = "CHT_GRP_3";
        public const string GCM_CHT_GRP_4 = "CHT_GRP_4";
        public const string GCM_CHT_GRP_5 = "CHT_GRP_5";
        public const string GCM_CHT_GRP_6 = "CHT_GRP_6";
        public const string GCM_CHT_GRP_7 = "CHT_GRP_7";
        public const string GCM_CHT_GRP_8 = "CHT_GRP_8";
        public const string GCM_CHT_GRP_9 = "CHT_GRP_9";
        public const string GCM_CHT_GRP_10 = "CHT_GRP_10";

        //SPC Chart Group Table 1~10
        public const string GCM_CHTSET_GRP_1 = "CHTSET_GRP_1";
        public const string GCM_CHTSET_GRP_2 = "CHTSET_GRP_2";
        public const string GCM_CHTSET_GRP_3 = "CHTSET_GRP_3";
        public const string GCM_CHTSET_GRP_4 = "CHTSET_GRP_4";
        public const string GCM_CHTSET_GRP_5 = "CHTSET_GRP_5";
        public const string GCM_CHTSET_GRP_6 = "CHTSET_GRP_6";
        public const string GCM_CHTSET_GRP_7 = "CHTSET_GRP_7";
        public const string GCM_CHTSET_GRP_8 = "CHTSET_GRP_8";
        public const string GCM_CHTSET_GRP_9 = "CHTSET_GRP_9";
        public const string GCM_CHTSET_GRP_10 = "CHTSET_GRP_10";

        public const string WIP_CREATE_CODE = "CREATE_CODE";
        public const string WIP_OWNER_CODE = "OWNER_CODE";
        public const string WIP_LOT_TYPE = "LOT_TYPE";
        public const string WIP_MATERIAL_TYPE = "MATERIAL_TYPE";
        public const string WIP_MATERIAL_PACKTYPE = "MATERIAL_PACK_TYPE";
        public const string WIP_OPTIONAL_FLOW_GROUP = "OPTIONAL_FLOW_GROUP";
        public const string WIP_OPTIONAL_OPER_GROUP = "OPTIONAL_OPER_GROUP";

        public const string WIP_UNIT_TABLE = "UNIT";
        public const string WIP_SHIP_CODE = "SHIP_CODE";
        public const string WIP_HOLD_CODE = "HOLD_CODE";
        public const string WIP_REPAIR_CODE = "REPAIR_CODE";
        public const string WIP_RESULT_CODE = "RESULT_CODE";
        public const string WIP_ACTION_CODE = "ACTION_CODE";
        public const string WIP_RELEASE_CODE = "RELEASE_CODE";
        public const string WIP_ORDER_STATUS = "ORDER_STATUS";
        public const string WIP_TERMINATE_CODE = "TERMINATE_CODE";
        public const string WIP_CV_CODE = "CV_CODE";

        public const string WIP_LOT_DEFECT_CODE = "LOT_DEFECT_CODE";

        public const string WIP_SUBLOT_GRADE = "SUBLOT_GRADE";

        public const string INV_SCRAP_CODE = "SCRAP_CODE";

        public const string RAS_RES_TYPE = "RES_TYPE";
        public const string RAS_SUBRES_TYPE = "SUBRES_TYPE";
        public const string RAS_AREA_CODE = "AREA";
        public const string RAS_SUBAREA_CODE = "SUB_AREA";
        public const string RAS_WORK_POSITION = "WORK_POSITION";
        public const string RAS_CHAMBER_GROUP = "CHAMBER_GROUP";
        public const string RAS_PM_PERIOD = "PM_PERIOD";
        public const string RAS_PM_EVENT = "PM_EVENT";

        public const string RAS_CRR_TYPE1 = "CRR_TYPE1";
        public const string RAS_CRR_TYPE2 = "CRR_TYPE2";
        public const string RAS_CRR_TYPE3 = "CRR_TYPE3";

        public const string ATTR_TYPE_TABLE = "ATTRIBUTE_TYPE";


        public const string RMA_CREATE_CODE = "RMA_CREATE_CODE";
        public const string RMA_RESULT_CODE = "RMA_RESULT_CODE";
        public const string ARCHIVE_MODULE = "ARCHIVE_MODULE";

        //POP
        public const string POP_PRINTER_TYPE = "PRINTER_TYPE";
        public const string POP_RESOLUTION = "RESOLUTION";
        public const string POP_TEXT_FONT = "TEXT_FONT";
        public const string POP_BARCODE_FONT = "BARCODE_FONT";
        public const string POP_PRINT_VARIABLE = "PRINT_VARIABLE";
        public const string POP_ROTATE = "ROTATE";
        public const string GCM_CUSTOMER = "CUSTOMER";

        //TOOL
        public const string RAS_TOOL_STATUS = "TOOL_STATUS";
        public const string RAS_TOOL_GRP = "TOOL_GRP";
        public const string RAS_TOOL_GRADE = "TOOL_GRADE";
        public const string RAS_TOOL_DEFECT = "TOOL_DEFECT";

        //PORT
        public const string RAS_PORT_STATE = "PORT_STATE";

        //QCM
        public const string QCM_SPROC_TYPE = "SPROC_TYPE";
        public const string QCM_INSP_METHOD = "QCM_INSP_METHOD";
        public const string QCM_DEFECT_GRP = "QCM_DEFECT_GRP";
        public const string QCM_INSP_TYPE = "QCM_INSP_TYPE";
        public const string QCM_VENDOR = "QCM_VENDOR";
        public const string QCM_CUSTOMER = "QCM_CUSTOMER";

        //WFM
        public const string WFM_NODE_TYPE = "WFM_NODE_TYPE";

        public const string GCM_TABLE_GROUP = "GCM_TABLE_GROUP";
        public const string SEC_FUNCTION_GROUP = "FUNCTION_GROUP";
        public const string SEC_PROGRAM_LIST = "PROGRAM_LIST";

        //EDC
        public const string EDC_UNIT_TABLE = "EDC_UNIT"; // Character Unit Table

        //SHEET
        public const string SHEET_EVENT = "SHEET_EVENT";

        public const string SHEET_QUERY_TYPE = "SHEET_QUERY_TYPE";
        public const string SHEET_SHEET_TYPE = "SHEET_SHEET_TYPE";
        public const string SHEET_DATA_TYPE = "SHEET_DATA_TYPE";

        public const string SHEET_TYPE_DEFINE = "SHEET_TYPE_DEFINE";
        public const string SHEET_TRAN_DEFINE = "SHEET_TRAN_DEFINE";

        public const string SHEET_GRP_CAPTION = "GROUP_CAPTION";
        public const string SHEET_GRP_TABLE = "GROUP_TABLE";

        public const string SHEET_TRN_CAPTION = "TRAN_CAPTION";
        public const string SHEET_TRN_TABLE = "TRAN_TABLE";
        //ID
        public const string ID_GEN_TYPE = "ID_GEN_TYPE";
        public const string ID_GEN_TRAN_CODE = "ID_GEN_TRAN_CODE";

        //Batch
        public const string BATCH_TYPE = "BATCH_TYPE";

        public const string FUNC_KEY_TYPE = "FUNC_KEY_TYPE";

        //Future Action
        public const string FAC_OA_SERVICES = "FAC_OA_SERVICES";

        public const string BIN_DATA_1 = "__BIN_DATA_1";
        public const string BIN_DATA_2 = "__BIN_DATA_2";
        public const string BIN_DATA_3 = "__BIN_DATA_3";
        public const string BIN_DATA_4 = "__BIN_DATA_4";
        public const string BIN_DATA_5 = "__BIN_DATA_5";
        public const string BIN_DATA_6 = "__BIN_DATA_6";
        public const string BIN_DATA_7 = "__BIN_DATA_7";
        public const string BIN_DATA_8 = "__BIN_DATA_8";
        public const string BIN_DATA_9 = "__BIN_DATA_9";
        public const string BIN_DATA_10 = "__BIN_DATA_10";

        //Flexible Group
        public const string SCREEN_GRP = "SCREEN_GROUP";
        #endregion

        #region "Resource Status"
        //Status
        public const string RESOURCE_STATUS_WAIT = "WAIT";
        public const string RESOURCE_STATUS_PROC = "PROC";
        public const string DEFAULT_RESOURCE_STATUS = "WAIT";
        #endregion

        #region "Resource Up/Down status"
        public const char RES_UP_FLAG = 'U';
        public const char RES_DOWN_FLAG = 'D';
        public const char DEFAULT_RES_UP_DOWN_FLAG = 'U';
        #endregion

        #region "Lot Status"
        public const string LOT_STATUS_WAIT = "WAIT";
        public const string LOT_STATUS_PROC = "PROC";
        public const string LOT_STATUS_RESV = "RESV";
        #endregion

        #region "Privilege Type"
        public const string PRV_TYPE_RES = "RESOURCE";
        public const string PRV_TYPE_OPER = "OPERATION";
        public const string PRV_TYPE_GCMTBL = "GCMTABLE";
        public const string PRV_TYPE_SERVICE = "SERVICE";
        #endregion

        #region "Print Font"
        //Type
        public const string POP_TYPE_TEXT = "T";
        public const string POP_TYPE_BARCODE = "B";
        public const string POP_TYPE_IMAGE = "I";
        public const string POP_TYPE_GRAPHIC = "G";

        //Rotate
        public const string POP_ROTATE_NORMAL = "N";
        public const string POP_ROTATE_ROTATED = "R";
        public const string POP_ROTATE_INVERTED = "I";
        public const string POP_ROTATE_BOTTOMUP = "B";

        //Barcode Font
        public const string POP_BAR_FONT_128 = "C";
        public const string POP_BAR_FONT_3OF9 = "3";
        public const string POP_BAR_FONT_PDF417 = "7";
        public const string POP_BAR_FONT_EAN_8 = "8";
        public const string POP_BAR_FONT_UPC_E = "9";
        public const string POP_BAR_FONT_EAN_13 = "E";
        public const string POP_BAR_FONT_UPC_A = "U";

        //Print Port
        public const string POP_PRINT_PORT_LPT = "LPT";
        public const string POP_PRINT_PORT_COM = "COM";

        // Communication Value.
        public const int STX = 0x2; // Start of Text
        public const int ETX = 0x3; // End of Text
        public const int CR = 0xD; // Carriage Return
        public const int LF = 0xA; // Line Feed
        public const int HT = 0x9; // Horizontal Tab


        // Print Status 愿??蹂??
        public static string s_PrtStatus; // ?꾨┛???곹깭媛?Check.
        public static bool b_InputSTX; // STX 媛?ㅼ뼱?붾뒗吏 ?좊Т.
        public const int i_BufFull = 50; // Receive Buffer Full Check Count.
        public const string s_TimeOut = "05"; // Print Status Check-Out Timeout

        public const double PRINT_200_DPM = 8; //Dots per Milimeter(200DPI)
        public const double PRINT_300_DPM = 12; //Dots per Milimeter(300DPI)

        #endregion

        #region "Prompt"
        public const char PROMPT_ASCII = 'A';
        public const char PROMPT_NUMBER = 'N';
        public const char PROMPT_FLOAT = 'F';
        #endregion

        #region "Alarm"

        //Alarm Type
        public const char ALM_NORMAL = 'N';
        public const char ALM_RESOURCE = 'R';
        public const char ALM_AUTO_GATHER = 'A';

        //Alarm Level
        public const char ALM_LEVEL_INFO = 'I';
        public const char ALM_LEVEL_WARN = 'W';
        public const char ALM_LEVEL_ERROR = 'E';

        //Alarm Transaction Point
        public const char ALM_TRAN_START = 'S';
        public const char ALM_TRAN_SPLIT = 'P';
        public const char ALM_TRAN_END = 'E';
        public const char ALM_TRAN_REWORK = 'R';


        #endregion

        #region "SUBLOT"
        public const string SUBLOT_GOOD_GRADE = "G";
        public const string SUBLOT_SCRAP_GRADE = "S";
        #endregion

        #region "MFO OPTION"
        public const string MAT_TYPE_RMATERIAL = "RM";
        public const string MAT_TYPE_COMPONENT = "CM";
        public const string MAT_TYPE_RETRIEVED = "RC";

        public const string MFO_EXT_LOSS_TBL_DEF = "EXT_LOSS_TBL_DEF";
        public const string MFO_EXT_BONUS_TBL_DEF = "EXT_BONUS_TBL_DEF";
        public const string MFO_EXT_LOT_DEFECT_TBL = "EXT_LOT_DEFECT_TBL";
        #endregion

        #region "ETC Constant"
        public const int EXCEL_MAX_COL = 255; //(EXCel?먯꽌 ?덉슜?섎뒗 理쒕? Column??-1)
        public const int MAX_SLOT_CNT = 1000; //理쒕? Slot 媛?닔
        public const int MAX_BATCH_CNT = 12; //Batch???ы븿?좎닔 ?덈뒗 Lot??理쒕? 媛쒖닔

        public const string SYS_DOCK_MENU = "SYS_MENU_MENU";
        public const string SYS_DOCK_FAVORITES = "SYS_MENU_FAVORITES";
        public const string SYS_DOCK_OPERATION = "SYS_MENU_OPERATION";
        public const string SYS_DOCK_RESOURCE = "SYS_MENU_RESOURCE";
        public const string SYS_DOCK_DISPATCHER = "SYS_MENU_DISPATCHER";
        public const string SYS_DOCK_WORKFLOW = "SYS_MENU_MODELER";
        public const string SYS_DOCK_BBS = "SYS_MENU_BBS";

        //Public Const FONT_NAME As String = "援대┝"
        //Public Const FONT_SIZE As Single = 9.0

        public const int SP_MAX_COLUMN_WIDTH = 500;
        public const int SP_MIN_COLUMN_WIDTH = 20;

        public const int LV_MAX_COLUMN_WIDTH = 500;
        public const int LV_MIN_COLUMN_WIDTH = 50;
        public const int LV_MAX_LIST_COUNT = 20;
        public const int LV_BONUS_LISTVIEW_HEIGHT = 2;
        public const int LV_BONUS_COLUMN_WIDTH = 4;
        public const int LV_BONUS_COLUMN_WIDTH_WITH_IMAGE = 2;

        public const int MDI_CHILD_HEIGHT = 580; //MDI Child Form???믪씠
        public const int MDI_CHILD_WIDTH = 750; //MDI Child Form????

        public const double MAX_QTY = 9999999.999; //Max Qty
        public const double MIN_QTY = -9999999.999; //Min Qty

        public const string DONT_CHECK_PASSWORD = "DO_NOT_CHECK_PASSWORD";

        public const int MAX_GDI_COUNT = 8000;

        #endregion
    }
}