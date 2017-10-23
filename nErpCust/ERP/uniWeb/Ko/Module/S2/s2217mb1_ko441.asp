
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 

    Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
    Call LoadBasisGlobalInf()

    Dim lgStrPrevKey
    On Error Resume Next                                                   '☜: Protect prorgram from crashing

    Err.Clear                                                              '☜: Clear Error status
    
    Call HideStatusWnd                                                     '☜: Hide Processing message

    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = UNICInt(Trim(Request("lgStrPrevKey")), 0)                   '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    If lgStrPrevKey = 0 Then

       lgStrSQL = "Select plant_cd,plant_nm " 
       lgStrSQL = lgStrSQL & " From B_Plant (Nolock) " 
       lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(lgKeyStream(0),"''", "S")

       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
          Response.Write  " <Script Language=vbscript>            " & vbCr
'hanc          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
'hanc          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
          Response.Write  "   Parent.Frm1.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
          Response.Write  " </Script> " & vbCr
          Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
          Call SubBizQueryMulti()
       Else 
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
       End If 
    Else 
       Call SubBizQueryMulti() 
    End If 
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount
    Dim dBillDt, dBillDtFirst
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

    dBillDt     = Request("txtBillFrDt")  '20071226::hanc
    dBillDtFirst     = left(dBillDt,8) & "01"  '20071226::hanc
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
'20071227::hanc
lgStrSQL = lgStrSQL & " SELECT NEP.PART,                                                                                                                     "
lgStrSQL = lgStrSQL & "        (CASE                                                                                                                         "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'A' THEN '매출금액'                                                                                        "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'B' THEN '예상매출'                                                                                        "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'B1' THEN '예상매출'                                                                                       "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'C' THEN '사업계획'                                                                                        "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'C1' THEN '사업계획'                                                                                       "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'D' THEN '전월실적'                                                                                        "
lgStrSQL = lgStrSQL & "         END) PART_NM,                                                                                                                "
lgStrSQL = lgStrSQL & "        (CASE                                                                                                                         "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'A' THEN '금액'                                                                                            "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'B' THEN '금액'                                                                                            "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'B1' THEN '달성율'                                                                                         "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'C' THEN '금액'                                                                                            "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'C1' THEN '달성율'                                                                                         "
lgStrSQL = lgStrSQL & "           WHEN NEP.PART = 'D' THEN '금액'                                                                                            "
lgStrSQL = lgStrSQL & "         END) PART_NM_SUB,                                                                                                            "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.TOD,0) AS TOD,                                                                                                     "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.W1,0) AS W1,                                                                                                       "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.W2,0) AS W2,                                                                                                       "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.W3,0) AS W3,                                                                                                       "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.W4,0) AS W4,                                                                                                       "
lgStrSQL = lgStrSQL & "        ISNULL(NEP.W5,0) AS W5,                                                                                                       "
lgStrSQL = lgStrSQL & "        cast(DATEPART(WEEK,'" & dBillDtFirst & "') as char(2)) +  '주 '+ dbo.ufn_GetWeek('" & dBillDtFirst & "', 1)  AS W1_NM,           "
lgStrSQL = lgStrSQL & "        cast(DATEPART(WEEK,'" & dBillDtFirst & "')+1 as char(2)) +  '주 '+ dbo.ufn_GetWeek('" & dBillDtFirst & "', 2)  AS W2_NM,         "
lgStrSQL = lgStrSQL & "        cast(DATEPART(WEEK,'" & dBillDtFirst & "')+2 as char(2)) +  '주 '+ dbo.ufn_GetWeek('" & dBillDtFirst & "', 3)  AS W3_NM,         "
lgStrSQL = lgStrSQL & "        cast(DATEPART(WEEK,'" & dBillDtFirst & "')+3 as char(2)) +  '주 '+ dbo.ufn_GetWeek('" & dBillDtFirst & "', 4)  AS W4_NM,         "
lgStrSQL = lgStrSQL & "        cast(DATEPART(WEEK,'" & dBillDtFirst & "')+4 as char(2)) +  '주 '+ dbo.ufn_GetWeek('" & dBillDtFirst & "', 5)  AS W5_NM          "
lgStrSQL = lgStrSQL & " FROM   (SELECT 'A' AS PART,                                                                                                          "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                             "
lgStrSQL = lgStrSQL & "                        WHERE  CONVERT(CHAR(10),S_BILL_HDR.BILL_DT,120) = '" & dBillDt & "'                                               "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP  LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                    "
lgStrSQL = lgStrSQL & "                       0) TOD,                                                                                                        "
lgStrSQL = lgStrSQL & "                (SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                        "
lgStrSQL = lgStrSQL & "                 FROM   S_BILL_HDR                                                                                                    "
lgStrSQL = lgStrSQL & "                 WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,'" & dBillDtFirst & "')                                               "
lgStrSQL = lgStrSQL & "                        AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ) W1,                                                                        "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                             "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,'" & dBillDtFirst & "') + 1                                    "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                    "
lgStrSQL = lgStrSQL & "                       0) W2,                                                                                                         "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                             "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,'" & dBillDtFirst & "') + 2                                    "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                    "
lgStrSQL = lgStrSQL & "                       0) W3,                                                                                                         "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                             "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,'" & dBillDtFirst & "') + 3                                    "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                    "
lgStrSQL = lgStrSQL & "                       0) W4,                                                                                                         "
lgStrSQL = lgStrSQL & "                (SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                        "
lgStrSQL = lgStrSQL & "                 FROM   S_BILL_HDR                                                                                                    "
lgStrSQL = lgStrSQL & "                 WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,'" & dBillDtFirst & "') + 4                                           "
lgStrSQL = lgStrSQL & "                        AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ) W5                                                                         "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                    "
lgStrSQL = lgStrSQL & "         SELECT 'B' AS PART,                                                                                                          "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') AS CHAR(2))             "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) / 7 AS TOD,                                                                                                 "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') AS CHAR(2))             "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W1,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 1 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W2,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 2 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W3,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 3 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W4,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'E'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 4 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W5                                                                                                       "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                    "
lgStrSQL = lgStrSQL & "         SELECT 'B1' AS PART,                                                                                                         "
lgStrSQL = lgStrSQL & "                0 TOD,                                                                                                                "
lgStrSQL = lgStrSQL & "                0 AS W1,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W2,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W3,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W4,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W5                                                                                                               "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                    "
lgStrSQL = lgStrSQL & "         SELECT 'C' AS PART,                                                                                                          "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') AS CHAR(2))             "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) / 7 AS TOD,                                                                                                 "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') AS CHAR(2))             "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W1,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 1 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W2,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 2 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W3,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 3 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W4,                                                                                                      "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_SP_ITEM_BY_BP.AMT_LOC,0))                                                                 "
lgStrSQL = lgStrSQL & "                        FROM   S_SP_ITEM_BY_BP                                                                                        "
lgStrSQL = lgStrSQL & "                        WHERE  S_SP_ITEM_BY_BP.SP_TYPE = 'M'                                                                          "
lgStrSQL = lgStrSQL & "                               AND SP_PERIOD = SUBSTRING('" & dBillDtFirst & "',1,4) + CAST(DATEPART(WEEK,'" & dBillDtFirst & "') + 4 AS CHAR(2))         "
lgStrSQL = lgStrSQL & "                               AND S_SP_ITEM_BY_BP.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                               "
lgStrSQL = lgStrSQL & "                       0) AS W5                                                                                                       "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                    "
lgStrSQL = lgStrSQL & "         SELECT 'C1' AS PART,                                                                                                         "
lgStrSQL = lgStrSQL & "                0 AS TOD,                                                                                                             "
lgStrSQL = lgStrSQL & "                0 AS W1,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W2,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W3,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W4,                                                                                                              "
lgStrSQL = lgStrSQL & "                0 AS W5                                                                                                               "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                    "
lgStrSQL = lgStrSQL & "         SELECT 'D' AS PART,                                                                                                                             "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                    "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                                                "
lgStrSQL = lgStrSQL & "                        WHERE  CONVERT(CHAR(10),S_BILL_HDR.BILL_DT,120) = convert(char(10), DATEADD(month, -1, '" & dBillDt & "'), 120)                       "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                                       "
lgStrSQL = lgStrSQL & "                       0) TOD,                                                                                                                           "
lgStrSQL = lgStrSQL & "                (SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                           "
lgStrSQL = lgStrSQL & "                 FROM   S_BILL_HDR                                                                                                                       "
lgStrSQL = lgStrSQL & "                 WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,convert(char(10), DATEADD(month, -1, '" & dBillDtFirst & "'), 120))                      "
lgStrSQL = lgStrSQL & "                        AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ) W1,                                                                                           "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                    "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                                                "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,convert(char(10), DATEADD(month, -1, '" & dBillDtFirst & "'), 120)) + 1           "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                                       "
lgStrSQL = lgStrSQL & "                       0) W2,                                                                                                                            "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                    "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                                                "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,convert(char(10), DATEADD(month, -1, '" & dBillDtFirst & "'), 120)) + 2           "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                                       "
lgStrSQL = lgStrSQL & "                       0) W3,                                                                                                                            "
lgStrSQL = lgStrSQL & "                ISNULL((SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                    "
lgStrSQL = lgStrSQL & "                        FROM   S_BILL_HDR                                                                                                                "
lgStrSQL = lgStrSQL & "                        WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,convert(char(10), DATEADD(month, -1, '" & dBillDtFirst & "'), 120)) + 3           "
lgStrSQL = lgStrSQL & "                               AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ),                                                                                       "
lgStrSQL = lgStrSQL & "                       0) W4,                                                                                                                            "
lgStrSQL = lgStrSQL & "                (SELECT SUM(ISNULL(S_BILL_HDR.BILL_AMT_LOC,0))                                                                                           "
lgStrSQL = lgStrSQL & "                 FROM   S_BILL_HDR                                                                                                                       "
lgStrSQL = lgStrSQL & "                 WHERE  DATEPART(WEEK,S_BILL_HDR.BILL_DT) = DATEPART(WEEK,convert(char(10), DATEADD(month, -1, '" & dBillDtFirst & "'), 120)) + 4                  "
lgStrSQL = lgStrSQL & "                        AND S_BILL_HDR.SALES_GRP LIKE " & FilterVar(Trim(Request("txtConSalesGrp")) & "%", " ", "S") & " ) W5                                                                                            "
lgStrSQL = lgStrSQL & " 			) NEP                                                                                                                                       "


'20071221::hanc    lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " d.wc_cd,d.wc_nm,a.item_cd,e.item_nm,e.spec,b.lot_no,b.lot_sub_no,b.good_on_hand_qty "
'20071221::hanc	lgStrSQL = lgStrSQL & " from b_item_by_plant a, "
'20071221::hanc	lgStrSQL = lgStrSQL & "         i_onhand_stock_detail b, "
'20071221::hanc	lgStrSQL = lgStrSQL & "         i_goods_movement_detail c, "
'20071221::hanc	lgStrSQL = lgStrSQL & "         p_work_center d, "
'20071221::hanc	lgStrSQL = lgStrSQL & "         b_item e	 "
'20071221::hanc	lgStrSQL = lgStrSQL & " where a.plant_cd =  " &  FilterVar(lgKeyStream(0),"''", "S")
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.plant_cd = b.plant_cd  "
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.procur_type = 'M'  "
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.lot_flg = 'Y' "
'20071221::hanc	lgStrSQL = lgStrSQL & " and b.lot_no <> '*'  "
'20071221::hanc	lgStrSQL = lgStrSQL & " and b.good_on_hand_qty > 0 "
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.item_cd = b.item_cd "
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.plant_cd = c.plant_cd "
'20071221::hanc	lgStrSQL = lgStrSQL & " and a.item_cd = c.item_cd "
'20071221::hanc'	lgStrSQL = lgStrSQL & " and c.trns_type = 'MR' "
'20071221::hanc'	lgStrSQL = lgStrSQL & " and a.plant_cd = d.plant_cd "
'20071221::hanc'	lgStrSQL = lgStrSQL & " and c.wc_cd = d.wc_cd "
'20071221::hanc	lgStrSQL = lgStrSQL & " and b.plant_cd = c.plant_cd  "
'20071221::hanc	lgStrSQL = lgStrSQL & " and b.item_cd = c.item_cd  "
'20071221::hanc'	lgStrSQL = lgStrSQL & " and b.lot_no = c.lot_no " 
'20071221::hanc'	lgStrSQL = lgStrSQL & " and b.lot_sub_no = c.lot_sub_no "
'20071221::hanc'	lgStrSQL = lgStrSQL & " and a.item_cd = e.item_cd "
'20071221::hanc	lgStrSQL = lgStrSQL & " order by d.wc_cd,b.lot_sub_no "
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Exit Sub 
    Else    
      
	   If CDbl(lgStrPrevKey) > 0 Then
		  lgObjRs.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgStrPrevKey)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   End If   
       iDx = 1		
       
       lgstrData = ""
       lgLngMaxRow       = CLng(Request("txtMaxRows"))

       Do While Not lgObjRs.EOF
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PART"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PART_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PART_NM_SUB"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TOD"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("W1"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("W2"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("W3"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("W4"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("W5"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W1_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W2_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W3_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W4_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W5_NM"))

'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_cd"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_sub_no"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("good_on_hand_qty"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)

          lgObjRs.MoveNext

          iDx =  iDx + 1
         If iDx > C_SHEETMAXROWS_D Then
			 lgStrPrevKey = lgStrPrevKey + 1
             Exit Do
         End If        
      Loop 
    End If
         
    If iDx <= C_SHEETMAXROWS_D Then
	    lgStrPrevKey = ""            
    End If            
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   
       
    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim itxtSpread
    Dim arrRowVal
    Dim arrColVal
    Dim lgErrorPos
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgErrorPos        = ""                                                           '☜: Set to space

    itxtSpread = Trim(Request("txtSpread"))
    
    If itxtSpread = "" Then
       Exit Sub
    End If   
    
	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data
	
    For iDx = 0 To UBound(arrRowVal,1) - 1
        arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C" :  Call SubBizSaveMultiCreate(arrColVal)                        '☜: Create
            Case "U" :  Call SubBizSaveMultiUpdate(arrColVal)                        '☜: Update
            Case "D" :  Call SubBizSaveMultiDelete(arrColVal)                        '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
    
    If lgErrorStatus = "YES" Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.SubSetErrPos(Trim(""" & lgErrorPos & """))" & vbCr
       Response.Write  " </Script>                  " & vbCr
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.DBSaveOk            " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "INSERT INTO student("
    lgStrSQL = lgStrSQL & " SchoolCD     , StudentID    ,"    '3
    lgStrSQL = lgStrSQL & " StudentNM    , Grade        ,"    '5
    lgStrSQL = lgStrSQL & " Phone        , ZipCd        ,"    '7
    lgStrSQL = lgStrSQL & " StudyOnOff   , EnrollDT     ,"    '9
    lgStrSQL = lgStrSQL & " GraduatedDT  , SMoney       ,"    '11
    lgStrSQL = lgStrSQL & " SMoneyCnt    , INSRT_UID    ,"    '13
    lgStrSQL = lgStrSQL & " INSRT_DT     , UPDT_UID     ,"    '15
    lgStrSQL = lgStrSQL & " UPDT_DT      )"    '16
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(02)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(03)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(04) ,"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(05) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(06) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(07) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(08) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(09)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(10)),"","S") & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(11),0)         & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(12),0)         & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate()," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate())" 
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE STUDENT SET "
    lgStrSQL = lgStrSQL & " StudentNM   = " & FilterVar(            arrColVal(04) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " Grade       = " & FilterVar(            arrColVal(05) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " Phone       = " & FilterVar(            arrColVal(06) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " ZipCd       = " & FilterVar(            arrColVal(07) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " StudyOnOff  = " & FilterVar(            arrColVal(08) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " EnrollDT    = " & FilterVar(UniConvDate(arrColVal(09)),Null,"S")  & ","
    lgStrSQL = lgStrSQL & " GraduatedDT = " & FilterVar(UniConvDate(arrColVal(10)),Null,"S")  & ","
    lgStrSQL = lgStrSQL & " SMoney      = " &            UNIConvNum(arrColVal(11),0)          & ","
    lgStrSQL = lgStrSQL & " SMoneyCnt   = " &            UNIConvNum(arrColVal(12),0)          & ","          
    lgStrSQL = lgStrSQL & " UPDT_UID    = " & FilterVar(gUsrId,"","S")                        & ","             
    lgStrSQL = lgStrSQL & " UPDT_DT     = GetDate() " 
    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "DELETE  FROM STUDENT"
    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>


