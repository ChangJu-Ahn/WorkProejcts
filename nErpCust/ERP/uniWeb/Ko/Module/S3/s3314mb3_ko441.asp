
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
        Case CStr(UID_M0004)                                                         '☜: 20080303::hanc 생산계획기간 가져오기
             Call SubBizQueryPeriod()
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
'    If lgStrPrevKey = 0 Then
'
'       lgStrSQL = "Select plant_cd,plant_nm " 
'       lgStrSQL = lgStrSQL & " From B_Plant (Nolock) " 
'       lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(lgKeyStream(0),"''", "S")
'
'       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
'          Response.Write  " <Script Language=vbscript>            " & vbCr
'          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
'          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
'          Response.Write  "   Parent.Frm1.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
'          Response.Write  " </Script> " & vbCr
'          Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
'          Call SubBizQueryMulti()
'       Else
'          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
'       End If
'    Else
       Call SubBizQueryMulti()
'    End If 
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

'20080303::hanc
Sub SubBizQueryPeriod()
    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount
    Dim tPeriod
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    lgStrSQL =  "SELECT reference                            "
    lgStrSQL = lgStrSQL & "FROM NEPES..B_CONFIGURATION       "
    lgStrSQL = lgStrSQL & "where major_cd like 'ZZ002'       "
    lgStrSQL = lgStrSQL & "and   minor_cd = 'DAYS'           "


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
          tPeriod = ConvSPChars(lgObjRs("reference"))


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
       Response.Write  " parent.frm1.txtPeriod.value  = """ & UCase(Trim(tPeriod)) & """" & vbCr      
'       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
'       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
'       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.InitSpreadSheet   " & vbCr           '20080303::hanc
       Response.Write  " </Script>             " & vbCr
    End If

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
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
	lgStrSQL = ""
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & "  declare @ProdPlanMonth DATETIME                                                                                                                                                  "
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " set @ProdPlanMonth = " & FilterVar(lgKeyStream(0), "''", "S")
'                    lgStrSQL = lgStrSQL & vbCrLf & "  set @ProdPlanMonth = '2007-12-01'                                                                                                                                                "
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " SELECT	DISTINCT A.ITEM_CD,                                                                                                                                                        "
    lgStrSQL = lgStrSQL & vbCrLf & " 		B.ITEM_NM,                                                                                                                                                                 "
    lgStrSQL = lgStrSQL & vbCrLf & " 		'' tracking_no,                                                                                                                                                            "
    lgStrSQL = lgStrSQL & vbCrLf & " 		'수주계획(수주일)' TYPE,                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 0, @ProdPlanMonth)) QTY01,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 1, @ProdPlanMonth)) QTY02,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 2, @ProdPlanMonth)) QTY03,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 3, @ProdPlanMonth)) QTY04,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 4, @ProdPlanMonth)) QTY05,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 5, @ProdPlanMonth)) QTY06,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 6, @ProdPlanMonth)) QTY07,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 7, @ProdPlanMonth)) QTY08,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 8, @ProdPlanMonth)) QTY09,     "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 9, @ProdPlanMonth)) QTY010,    "
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 10, @ProdPlanMonth)) QTY011,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 11, @ProdPlanMonth)) QTY012,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 12, @ProdPlanMonth)) QTY013,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 13, @ProdPlanMonth)) QTY014,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 14, @ProdPlanMonth)) QTY015,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 15, @ProdPlanMonth)) QTY016,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 16, @ProdPlanMonth)) QTY017,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 17, @ProdPlanMonth)) QTY018,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 18, @ProdPlanMonth)) QTY019,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 19, @ProdPlanMonth)) QTY020,   "
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 20, @ProdPlanMonth)) QTY021,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 21, @ProdPlanMonth)) QTY022,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 22, @ProdPlanMonth)) QTY023,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 23, @ProdPlanMonth)) QTY024,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 24, @ProdPlanMonth)) QTY025,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 25, @ProdPlanMonth)) QTY026,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 26, @ProdPlanMonth)) QTY027,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 27, @ProdPlanMonth)) QTY028,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 28, @ProdPlanMonth)) QTY029,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 29, @ProdPlanMonth)) QTY030,   "
    lgStrSQL = lgStrSQL & vbCrLf & "                                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 30, @ProdPlanMonth)) QTY031,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 31, @ProdPlanMonth)) QTY032,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 32, @ProdPlanMonth)) QTY033,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 33, @ProdPlanMonth)) QTY034,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 34, @ProdPlanMonth)) QTY035,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 35, @ProdPlanMonth)) QTY036,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 36, @ProdPlanMonth)) QTY037,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 37, @ProdPlanMonth)) QTY038,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 38, @ProdPlanMonth)) QTY039,   "
    lgStrSQL = lgStrSQL & vbCrLf & " 		(SELECT PLAN_QTY FROM SALES_ITEM_REQ_PLAN_KO441 WHERE PROJECT_CODE = A.PROJECT_CODE AND ITEM_CD = A.ITEM_CD AND DLVY_PLAN_DT = DATEADD(day, 39, @ProdPlanMonth)) QTY040,   "
    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_1,                                                                                                                                                                "
    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_2,                                                                                                                                                                "
    lgStrSQL = lgStrSQL & vbCrLf & "         0 MONTH_3                                                                                                                                                                 "
    lgStrSQL = lgStrSQL & vbCrLf & " FROM	SALES_ITEM_REQ_PLAN_KO441 A,                                                                                                                                               "
    lgStrSQL = lgStrSQL & vbCrLf & " 		B_ITEM B                                                                                                                                                                   "
    lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.ITEM_CD = B.ITEM_CD                                                                                                                                                      "

'    lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " d.wc_cd,d.wc_nm,a.item_cd,e.item_nm,e.spec,b.lot_no,b.lot_sub_no,b.good_on_hand_qty "
'	lgStrSQL = lgStrSQL & " from b_item_by_plant a, "
'	lgStrSQL = lgStrSQL & "         i_onhand_stock_detail b, "
'	lgStrSQL = lgStrSQL & "         i_goods_movement_detail c, "
'	lgStrSQL = lgStrSQL & "         p_work_center d, "
'	lgStrSQL = lgStrSQL & "         b_item e	 "
'	lgStrSQL = lgStrSQL & " where a.plant_cd =  " &  FilterVar(lgKeyStream(0),"''", "S")
'	lgStrSQL = lgStrSQL & " and a.plant_cd = b.plant_cd  "
'	lgStrSQL = lgStrSQL & " and a.procur_type = 'M'  "
'	lgStrSQL = lgStrSQL & " and a.lot_flg = 'Y' "
'	lgStrSQL = lgStrSQL & " and b.lot_no <> '*'  "
'	lgStrSQL = lgStrSQL & " and b.good_on_hand_qty > 0 "
'	lgStrSQL = lgStrSQL & " and a.item_cd = b.item_cd "
'	lgStrSQL = lgStrSQL & " and a.plant_cd = c.plant_cd "
'	lgStrSQL = lgStrSQL & " and a.item_cd = c.item_cd "
''	lgStrSQL = lgStrSQL & " and c.trns_type = 'MR' "
''	lgStrSQL = lgStrSQL & " and a.plant_cd = d.plant_cd "
''	lgStrSQL = lgStrSQL & " and c.wc_cd = d.wc_cd "
'	lgStrSQL = lgStrSQL & " and b.plant_cd = c.plant_cd  "
'	lgStrSQL = lgStrSQL & " and b.item_cd = c.item_cd  "
''	lgStrSQL = lgStrSQL & " and b.lot_no = c.lot_no " 
''	lgStrSQL = lgStrSQL & " and b.lot_sub_no = c.lot_sub_no "
''	lgStrSQL = lgStrSQL & " and a.item_cd = e.item_cd "
'	lgStrSQL = lgStrSQL & " order by d.wc_cd,b.lot_sub_no "
        
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
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_cd"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_sub_no"))
'          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("good_on_hand_qty"), ggQty.DecPoint, 0)

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tracking_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("type"))
            
'                If UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) = 0 Then
'                lgstrData = lgstrData & Chr(11) & ""
'                Else
'                lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
'                End If
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty02"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty03"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty04"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty05"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty06"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty07"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty08"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty09"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty10"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty11"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty12"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty13"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty14"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty15"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty16"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty17"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty18"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty19"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty20"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty21"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty22"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty23"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty24"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty25"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty26"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty27"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty28"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty29"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty30"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty31"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty32"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::32
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty33"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::33
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty34"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::34
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty35"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::35
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty36"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::36

            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty37"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::37
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty38"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::38
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty39"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::39
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty40"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)      '20080303::hanc::40
            
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_1"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_2"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("month_3"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & ""	'ConvSPChars(lgObjRs("plant_cd"))
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty01"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty02"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty03"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty04"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty05"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty06"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty07"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty08"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty09"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty10"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty11"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty12"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty13"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty14"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty15"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty16"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty17"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty18"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty19"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty20"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty21"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty22"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty23"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty24"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty25"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty26"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty27"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty28"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty29"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty30"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("qty31"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)

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

    Call SubBizDelBeforeCreate()	'20080303::hanc

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

'20080312::hanc::begin-------------------------------------------------------
' 엑셀에 정의된 item이 중복되었을 경우 Stop
'----------------------------------------------------------------------------
    Dim iCount
        
        lgStrSQL = "Select  count(*) cnt  " 
        lgStrSQL = lgStrSQL & " From sales_Item_Req_Plan_terminal_Ko441 (Nolock) " 
        lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(      UCase(arrColVal(02)),"","S") & " "
        lgStrSQL = lgStrSQL & " AND ITEM_cd     = " & FilterVar(      UCase(arrColVal(03)),"","S") & " "
        lgStrSQL = lgStrSQL & " AND sales_group     = " & FilterVar(      UCase(arrColVal(86)),"","S") & " "
        
        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
            iCount = UniConvNumberDBToCompany(lgObjRs("cnt"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
        End If
        
        if iCount > 0 then
            Call DisplayMsgBox("ZZ0013", vbInformation, FilterVar(      UCase(arrColVal(03)),"","S"), "", I_MKSCRIPT)                  '☜: No data is found. 
            lgStrPrevKey  = ""
            lgErrorStatus = "YES"
            Exit Sub 
        end if  
'20080312::hanc::end  -------------------------------------------------------

'20080312::hanc::begin-------------------------------------------------------
' 엑셀에 정의된 item이 b_item_by_plant에 없는 품목일 경우 Stop (ex. 오타 ...)
'----------------------------------------------------------------------------
    Dim iCount1
        
        lgStrSQL = "Select  count(*) cnt  " 
        lgStrSQL = lgStrSQL & " From b_item_by_plant (Nolock) " 
        lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(      UCase(arrColVal(02)),"","S") & " "
        lgStrSQL = lgStrSQL & " AND ITEM_cd     = " & FilterVar(      UCase(arrColVal(03)),"","S") & " "
        
        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
            iCount1 = UniConvNumberDBToCompany(lgObjRs("cnt"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
        End If
        
        if iCount1 = 0 then
            Call DisplayMsgBox("ZZ0014", vbInformation, FilterVar(      UCase(arrColVal(03)),"","S"), "", I_MKSCRIPT)                  '☜: No data is found. 
            lgStrPrevKey  = ""
            lgErrorStatus = "YES"
            Exit Sub 
        end if  
'20080312::hanc::end  -------------------------------------------------------

    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " INSERT INTO sales_Item_Req_Plan_terminal_Ko441       "
    lgStrSQL = lgStrSQL & "(                                          "
    lgStrSQL = lgStrSQL & "    PLANT_CD,                              "
    lgStrSQL = lgStrSQL & "    sales_group,                               "
    lgStrSQL = lgStrSQL & "    ITEM_CD,                               "
    lgStrSQL = lgStrSQL & "    DAY_CNT,                               "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT1,                         "                     '1
    lgStrSQL = lgStrSQL & "    PLAN_QTY1,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT2,                         "                     '2
    lgStrSQL = lgStrSQL & "    PLAN_QTY2,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT3,                         "                     '3
    lgStrSQL = lgStrSQL & "    PLAN_QTY3,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT4,                         "                     '4
    lgStrSQL = lgStrSQL & "    PLAN_QTY4,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT5,                         "                     '5
    lgStrSQL = lgStrSQL & "    PLAN_QTY5,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT6,                         "                     '6
    lgStrSQL = lgStrSQL & "    PLAN_QTY6,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT7,                         "                     '7
    lgStrSQL = lgStrSQL & "    PLAN_QTY7,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT8,                         "                     '8
    lgStrSQL = lgStrSQL & "    PLAN_QTY8,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT9,                         "                     '9
    lgStrSQL = lgStrSQL & "    PLAN_QTY9,                             "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT10,                        "                     '10
    lgStrSQL = lgStrSQL & "    PLAN_QTY10,                            "

    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT11,                        "                     '11
    lgStrSQL = lgStrSQL & "    PLAN_QTY11,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT12,                        "                     '12
    lgStrSQL = lgStrSQL & "    PLAN_QTY12,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT13,                        "                     '13
    lgStrSQL = lgStrSQL & "    PLAN_QTY13,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT14,                        "                     '14
    lgStrSQL = lgStrSQL & "    PLAN_QTY14,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT15,                        "                     '15
    lgStrSQL = lgStrSQL & "    PLAN_QTY15,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT16,                        "                     '16
    lgStrSQL = lgStrSQL & "    PLAN_QTY16,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT17,                        "                     '17
    lgStrSQL = lgStrSQL & "    PLAN_QTY17,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT18,                        "                     '18
    lgStrSQL = lgStrSQL & "    PLAN_QTY18,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT19,                        "                     '19
    lgStrSQL = lgStrSQL & "    PLAN_QTY19,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT20,                        "                     '20
    lgStrSQL = lgStrSQL & "    PLAN_QTY20,                            "

    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT21,                        "                     '21
    lgStrSQL = lgStrSQL & "    PLAN_QTY21,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT22,                        "                     '22
    lgStrSQL = lgStrSQL & "    PLAN_QTY22,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT23,                        "                     '23
    lgStrSQL = lgStrSQL & "    PLAN_QTY23,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT24,                        "                     '24
    lgStrSQL = lgStrSQL & "    PLAN_QTY24,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT25,                        "                     '25
    lgStrSQL = lgStrSQL & "    PLAN_QTY25,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT26,                        "                     '26
    lgStrSQL = lgStrSQL & "    PLAN_QTY26,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT27,                        "                     '27
    lgStrSQL = lgStrSQL & "    PLAN_QTY27,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT28,                        "                     '28
    lgStrSQL = lgStrSQL & "    PLAN_QTY28,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT29,                        "                     '29
    lgStrSQL = lgStrSQL & "    PLAN_QTY29,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT30,                        "                     '30
    lgStrSQL = lgStrSQL & "    PLAN_QTY30,                            "

    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT31,                        "                     '31
    lgStrSQL = lgStrSQL & "    PLAN_QTY31,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT32,                        "                     '32
    lgStrSQL = lgStrSQL & "    PLAN_QTY32,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT33,                        "                     '33
    lgStrSQL = lgStrSQL & "    PLAN_QTY33,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT34,                        "                     '34
    lgStrSQL = lgStrSQL & "    PLAN_QTY34,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT35,                        "                     '35
    lgStrSQL = lgStrSQL & "    PLAN_QTY35,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT36,                        "                     '36
    lgStrSQL = lgStrSQL & "    PLAN_QTY36,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT37,                        "                     '37
    lgStrSQL = lgStrSQL & "    PLAN_QTY37,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT38,                        "                     '38
    lgStrSQL = lgStrSQL & "    PLAN_QTY38,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT39,                        "                     '39
    lgStrSQL = lgStrSQL & "    PLAN_QTY39,                            "
    lgStrSQL = lgStrSQL & "    DLVY_PLAN_DT40,                        "                     '40
    lgStrSQL = lgStrSQL & "    PLAN_QTY40,                             "
    lgStrSQL = lgStrSQL & "    insrt_user_id    ,"
    lgStrSQL = lgStrSQL & "    insrt_dt     "
    lgStrSQL = lgStrSQL & ")                                          "
    lgStrSQL = lgStrSQL & "VALUES                                     "
    lgStrSQL = lgStrSQL & "(                                          "
    lgStrSQL = lgStrSQL &      FilterVar(      UCase(arrColVal(02)),"","S") & ","
    lgStrSQL = lgStrSQL &      FilterVar(      UCase(arrColVal(86)),"","S") & ","
    lgStrSQL = lgStrSQL &      FilterVar(      UCase(arrColVal(03)),"","S") & ","
    lgStrSQL = lgStrSQL &     UNIConvNum (arrColVal(05),0)         & ","

    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(06)),"","S") & ","           '1
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(07),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(08)),"","S") & ","           '2
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(09),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(10)),"","S") & ","           '3
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(11),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(12)),"","S") & ","           '4
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(13),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(14)),"","S") & ","           '5
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(15),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(16)),"","S") & ","           '6
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(17),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(18)),"","S") & ","           '7
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(19),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(20)),"","S") & ","           '8
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(21),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(22)),"","S") & ","           '9
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(23),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(24)),"","S") & ","           '10
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(25),0)         & ", "

    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(26)),"","S") & ","           '11
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(27),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(28)),"","S") & ","           '12
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(29),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(30)),"","S") & ","           '13
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(31),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(32)),"","S") & ","           '14
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(33),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(34)),"","S") & ","           '15
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(35),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(36)),"","S") & ","           '16
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(37),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(38)),"","S") & ","           '17
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(39),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(40)),"","S") & ","           '18
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(41),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(42)),"","S") & ","           '19
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(43),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(44)),"","S") & ","           '20
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(45),0)         & ", "

    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(46)),"","S") & ","           '21
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(47),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(48)),"","S") & ","           '22
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(49),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(50)),"","S") & ","           '23
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(51),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(52)),"","S") & ","           '24
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(53),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(54)),"","S") & ","           '25
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(55),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(56)),"","S") & ","           '26
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(57),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(58)),"","S") & ","           '27
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(59),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(60)),"","S") & ","           '28
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(61),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(62)),"","S") & ","           '29
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(63),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(64)),"","S") & ","           '30
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(65),0)         & ", "

    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(66)),"","S") & ","           '31
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(67),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(68)),"","S") & ","           '32
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(69),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(70)),"","S") & ","           '33
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(71),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(72)),"","S") & ","           '34
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(73),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(74)),"","S") & ","           '35
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(75),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(76)),"","S") & ","           '36
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(77),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(78)),"","S") & ","           '37
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(79),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(80)),"","S") & ","           '38
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(81),0)         & ", "                   
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(82)),"","S") & ","           '39
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(83),0)         & ", "
    lgStrSQL = lgStrSQL &      FilterVar(UniConvDate(arrColVal(84)),"","S") & ","           '40
    lgStrSQL = lgStrSQL &      UNIConvNum (arrColVal(85),0)         & ", "

    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate()                                        " 

    lgStrSQL = lgStrSQL & ")                                          "

'    lgStrSQL = "INSERT INTO sales_Item_Req_Plan_terminal_Ko441("
'    lgStrSQL = lgStrSQL & " SchoolCD     , StudentID    ,"    '3
'    lgStrSQL = lgStrSQL & " StudentNM    , Grade        ,"    '5
'    lgStrSQL = lgStrSQL & " Phone        , ZipCd        ,"    '7
'    lgStrSQL = lgStrSQL & " StudyOnOff   , EnrollDT     ,"    '9
'    lgStrSQL = lgStrSQL & " GraduatedDT  , SMoney       ,"    '11
'    lgStrSQL = lgStrSQL & " SMoneyCnt    , INSRT_UID    ,"    '13
'    lgStrSQL = lgStrSQL & " INSRT_DT     , UPDT_UID     ,"    '15
'    lgStrSQL = lgStrSQL & " UPDT_DT      )"    '16
'    lgStrSQL = lgStrSQL & " VALUES(" 
'    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(02)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(03)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(04) ,"","S") & "," 
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(05) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(06) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(07) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(08) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(09)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(10)),"","S") & ","
'    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(11),0)         & ","
'    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(12),0)         & ","
'    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
'    lgStrSQL = lgStrSQL & " GetDate()," 
'    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
'    lgStrSQL = lgStrSQL & " GetDate())" 
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'20080303::HANC------------------------------------------------------------------------------------------------------------------------
'INSERT 전에 sales_Item_Req_Plan_terminal_Ko441 의 모든 DATA를 지운다.
' sales_Item_Req_Plan_terminal_Ko441 테이블은 sales_Item_Req_Plan_Ko441 에 등록 직전의 통로테이블이다. 그렇다고 템프테이블이 아닌 정식 테이블이니 주의하도록한다.
' 이렇게 통로 테이블 사용한 이유는 엑셀양식을 바로 sales_Item_Req_Plan_Ko441 테이블에 넣으려니 해답이 나오지 않았기에 통로테이블에 넣고
' 트리거를 이용하여 sales_Item_Req_Plan_Ko441에 밀어 넣는 방식을 취하였다.
'---------------------------------------------------------------------------------------------------------------------------------------
Sub SubBizDelBeforeCreate()

    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " DELETE FROM  sales_Item_Req_Plan_terminal_Ko441       "
    
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

    Dim i
    Dim plantCd, itemCd, TrackingNo, mpsDt, mpsQty, mpsType, entUser, s_sales_group

    plantCd = FilterVar(UCase(arrColVal(2)), "''", "S")
    itemCd = FilterVar(UCase(arrColVal(3)), "''", "S")
    s_sales_group = FilterVar(UCase(arrColVal(4)), "''", "S")
'    TrackingNo = FilterVar(UCase(arrColVal(4)), "''", "S")
    entUser = FilterVar(gUsrId, "''", "S")

    For i = 0 To 39  '20080304::hanc:: 30
    
        If UNIConvNum(arrColVal(3 * i + 7), 0) <> UNIConvNum(arrColVal(3 * i + 8), 0) Then
			If len(Replace(arrColVal(3 * i + 6), "-", "")) < 2 Then
				strdt = "0" + Replace(arrColVal(3 * i + 6), "-", "")
			End If
            
            mpsDt = FilterVar(Replace(arrColVal(3 * i + 6), "-", ""), "''", "S")
            mpsQty = UNIConvNum(arrColVal(3 * i + 7), 0)

	        Call SubBizSaveMultiUpdateReal(plantCd, itemCd, mpsDt, mpsQty, entUser, s_sales_group)
			
			Call SubHandleError("MR", lgObjConn, lgObjRs, Err)
			Call SubCloseRs(lgObjRs)
			'----------------------------------------------------------------------------------------------------
        End If
    Next

'    '---------- Developer Coding part (Start) ---------------------------------------------------------------
'    'A developer must define field to update record
'    '--------------------------------------------------------------------------------------------------------
'    lgStrSQL = "UPDATE STUDENT SET "
'    lgStrSQL = lgStrSQL & " StudentNM   = " & FilterVar(            arrColVal(04) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " Grade       = " & FilterVar(            arrColVal(05) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " Phone       = " & FilterVar(            arrColVal(06) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " ZipCd       = " & FilterVar(            arrColVal(07) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " StudyOnOff  = " & FilterVar(            arrColVal(08) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " EnrollDT    = " & FilterVar(UniConvDate(arrColVal(09)),Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " GraduatedDT = " & FilterVar(UniConvDate(arrColVal(10)),Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " SMoney      = " &            UNIConvNum(arrColVal(11),0)          & ","
'    lgStrSQL = lgStrSQL & " SMoneyCnt   = " &            UNIConvNum(arrColVal(12),0)          & ","          
'    lgStrSQL = lgStrSQL & " UPDT_UID    = " & FilterVar(gUsrId,"","S")                        & ","             
'    lgStrSQL = lgStrSQL & " UPDT_DT     = GetDate() " 
'    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
'    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
'    
'    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
'
'    If CheckSYSTEMError(Err,True) = True Then
'       lgErrorStatus    = "YES"
'       ObjectContext.SetAbort
'    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdateReal(plantCd, itemCd, mpsDt, mpsQty, entUser, s_sales_group)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = " EXEC usp_s3314_ko441 " & plantCd & "," & itemCd & "," & mpsDt & "," & mpsQty & "," & entUser & "," & s_sales_group

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

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

    lgStrSQL = "DELETE  FROM sales_Item_Req_Plan_Ko441"
    lgStrSQL = lgStrSQL & " WHERE PROJECT_CODE  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND   ITEM_CD = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    lgStrSQL = lgStrSQL & " AND   sales_group = " &  FilterVar(Trim(UCase(arrColVal(6))),"''","S")

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


