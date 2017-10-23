
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
          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
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
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
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

    lgStrSQL = "          SELECT	A.PLANT_CD,                    "
    lgStrSQL = lgStrSQL & " 		A.PRODT_ORDER_NO_S,            "
    lgStrSQL = lgStrSQL & " 		[DBO].[ufn_GetItemCd](MES_ITEM_CD_S) MES_ITEM_CD_S,               "
    lgStrSQL = lgStrSQL & " 		A.PLANT_CD_TARGET,             "
    lgStrSQL = lgStrSQL & " 		A.PRODT_ORDER_NO_T,            "
    lgStrSQL = lgStrSQL & " 		[DBO].[ufn_GetItemCd](MES_ITEM_CD_T) MES_ITEM_CD_T,               "
    lgStrSQL = lgStrSQL & " 		A.JOB_QTY_T,                   "
    lgStrSQL = lgStrSQL & " 		A.CREATE_TYPE,                 "
    lgStrSQL = lgStrSQL & " 		A.SEND_DT,                     "
    lgStrSQL = lgStrSQL & " 		A.ERP_RECEIVE_DT,              "
    lgStrSQL = lgStrSQL & " 		A.ERP_APPLY_FLAG,              "
    lgStrSQL = lgStrSQL & " 		A.ERR_DESC                     "
    lgStrSQL = lgStrSQL & " FROM	T_IF_RCV_DEVICE_CHG_KO441 A    "
    lgStrSQL = lgStrSQL & " WHERE	A.PLANT_CD		=	" & FilterVar(Trim(Request("txtPlantCd")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	    [DBO].[ufn_GetItemCd](MES_ITEM_CD_S) like	" & FilterVar(Trim(Request("txtItemCds"))&"%" ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	    [DBO].[ufn_GetItemCd](MES_ITEM_CD_T) like	" & FilterVar(Trim(Request("txtItemCdt"))&"%" ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	CONVERT(CHAR(10),a.send_dt,120)		>=	" & FilterVar(Trim(Request("txtSendStartDt")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	CONVERT(CHAR(10),a.send_dt,120)		<=	" & FilterVar(Trim(Request("txtSendEndDt")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	CONVERT(CHAR(10),a.JOB_PLAN_DT_T,120)		>=	" & FilterVar(Trim(Request("txtPlanStartDt")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	CONVERT(CHAR(10),a.JOB_PLAN_DT_T,120)		<=	" & FilterVar(Trim(Request("txtPlanEndDt")) ,"","S")  & " "
    'lgStrSQL = lgStrSQL & " AND	A.ERP_APPLY_FLAG LIKE	" & FilterVar(Trim(Request("rdoFlag"))&"%" ,"","S")  & " "
    lgStrSQL = lgStrSQL & " AND	ISNULL(A.ERP_FLAG,'') LIKE	" & FilterVar(Trim(Request("rdoFlagF"))&"%" ,"","S")  & " "
    lgStrSQL = lgStrSQL & " ORDER BY	A.SEND_DT    "

        
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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRODT_ORDER_NO_S"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MES_ITEM_CD_S"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD_TARGET"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRODT_ORDER_NO_T"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MES_ITEM_CD_T"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("JOB_QTY_T"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_TYPE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEND_DT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ERP_RECEIVE_DT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ERP_APPLY_FLAG"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ERR_DESC"))
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


