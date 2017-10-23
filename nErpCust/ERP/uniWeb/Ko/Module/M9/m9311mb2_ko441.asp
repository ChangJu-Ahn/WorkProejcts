
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 
	Dim prefix
	Dim gl_dt
	Dim strdp_no, IntRetCD
	
	
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

    gl_dt   = Request("txtcurr_dt") 
    prefix  = "PQ"
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

        
'Call ServerMesgBox("PLEASE "  , vbInformation, I_MKSCRIPT)

    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
		Case "Release","UnRelease"			    '☜: 확정,확정취소 요청을 받음 
			 Call SubRelease()
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
        
    lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " d.wc_cd,d.wc_nm,a.item_cd,e.item_nm,e.spec,b.lot_no,b.lot_sub_no,b.good_on_hand_qty "
	lgStrSQL = lgStrSQL & " from b_item_by_plant a, "
	lgStrSQL = lgStrSQL & "         i_onhand_stock_detail b, "
	lgStrSQL = lgStrSQL & "         i_goods_movement_detail c, "
	lgStrSQL = lgStrSQL & "         p_work_center d, "
	lgStrSQL = lgStrSQL & "         b_item e	 "
	lgStrSQL = lgStrSQL & " where a.plant_cd =  " &  FilterVar(lgKeyStream(0),"''", "S")
	lgStrSQL = lgStrSQL & " and a.plant_cd = b.plant_cd  "
	lgStrSQL = lgStrSQL & " and a.procur_type = 'M'  "
	lgStrSQL = lgStrSQL & " and a.lot_flg = 'Y' "
	lgStrSQL = lgStrSQL & " and b.lot_no <> '*'  "
	lgStrSQL = lgStrSQL & " and b.good_on_hand_qty > 0 "
	lgStrSQL = lgStrSQL & " and a.item_cd = b.item_cd "
	lgStrSQL = lgStrSQL & " and a.plant_cd = c.plant_cd "
	lgStrSQL = lgStrSQL & " and a.item_cd = c.item_cd "
'	lgStrSQL = lgStrSQL & " and c.trns_type = 'MR' "
'	lgStrSQL = lgStrSQL & " and a.plant_cd = d.plant_cd "
'	lgStrSQL = lgStrSQL & " and c.wc_cd = d.wc_cd "
	lgStrSQL = lgStrSQL & " and b.plant_cd = c.plant_cd  "
	lgStrSQL = lgStrSQL & " and b.item_cd = c.item_cd  "
'	lgStrSQL = lgStrSQL & " and b.lot_no = c.lot_no " 
'	lgStrSQL = lgStrSQL & " and b.lot_sub_no = c.lot_sub_no "
'	lgStrSQL = lgStrSQL & " and a.item_cd = e.item_cd "
	lgStrSQL = lgStrSQL & " order by d.wc_cd,b.lot_sub_no "
        
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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_sub_no"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("good_on_hand_qty"), ggQty.DecPoint, 0)
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
'       Response.Write  " parent.frm1.txtPoNo1.value = '" & UCase(Trim(strdp_no)) & "'" & vbCr      
'       Response.Write  " parent.frm1.txtPoNo.value = '" & UCase(Trim(strdp_no)) & "'" & vbCr      
       Response.Write  " Parent.DBSaveOk111            " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim lgStrSQL
    Dim L_CNT
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    
    If Not IsNull(Trim(Request("txtPoNo1"))) And Trim(Request("txtPoNo1")) <> "" Then
    	strdp_no = Trim(Request("txtPoNo1"))
	Else
		Call SubAuto_number()
	End If
	
	If IsNull(strdp_no) Or strdp_no = "" Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()		
	End If
    
    lgStrSQL = " SELECT	COUNT(*) CNT                                                        "
    lgStrSQL = lgStrSQL & "FROM	M_ISSUE_REQ_HDR_KO441 A                                 "
    lgStrSQL = lgStrSQL & "WHERE	A.PLANT_CD		=	" & FilterVar(Trim(Request("txtPlantCd")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & "AND		A.ISSUE_REQ_NO	=	" & FilterVar(Trim(strdp_no) ,"","S")  & " "

    If FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else
        lgstrData = ""
        iDx = 1
        Do While Not lgObjRs.EOF
            L_CNT =  CInt(ConvSPChars(lgObjRs("CNT")))
            lgObjRs.MoveNext
        Loop
    End If
'Call ServerMesgBox("PLEASE 301"  , vbInformation, I_MKSCRIPT)

    IF L_CNT > 0 THEN       'update
        lgStrSQL =  "UPDATE  M_ISSUE_REQ_HDR_KO441 "
        lgStrSQL = lgStrSQL & "SET     REQ_DT           =    " & FilterVar(UCase(Trim(Request("txtPoDt"))) ,"","S")   & ", "        '2008-03-21 3:15오후 :: hanc
        lgStrSQL = lgStrSQL & "        TRNS_TYPE        =    " & FilterVar(UCase(Trim(Request("rdoReleaseflg"))) ,"","S")   & ", "
        lgStrSQL = lgStrSQL & "        ISSUE_TYPE       =    " & FilterVar(UCase(Trim(Request("txtPoTypeCd"))) ,"","S")     & ", "
        lgStrSQL = lgStrSQL & "        MOV_TYPE         =    " & FilterVar(UCase(Trim(Request("txtPoTypeCd"))) ,"","S")     & ", "
        lgStrSQL = lgStrSQL & "        EMP_NO           =    " & FilterVar(UCase(Trim(Request("txtEmp_no"))) ,"","S")       & ", "
        lgStrSQL = lgStrSQL & "        DEPT_CD          =    " & FilterVar(UCase(Trim(Request("txtDept_cd"))) ,"","S")      & ", "
        lgStrSQL = lgStrSQL & "        CONFIRM_FLAG     =    'N',               "    
        lgStrSQL = lgStrSQL & "        REMARK           =    " & FilterVar(UCase(Trim(Request("txtRemark"))) ,"","S")      & ", "
        lgStrSQL = lgStrSQL & "        INSRT_USER_ID    =    " & FilterVar(gUsrId,"","S")                          & "," 
        lgStrSQL = lgStrSQL & "        INSRT_DT         =    GETDATE(),         "    
        lgStrSQL = lgStrSQL & "        UPDT_USER_ID     =    " & FilterVar(gUsrId,"","S")                          & "," 
        lgStrSQL = lgStrSQL & "        UPDT_DT          =    GETDATE(),          "    
        lgStrSQL = lgStrSQL & "        LOC           =    " & FilterVar(UCase(Trim(Request("txtLoc"))) ,"","S")      & " "
        lgStrSQL = lgStrSQL & "WHERE	PLANT_CD		=	" & FilterVar(UCase(Trim(Request("txtPlantCd"))) ,"","S")  & " "
        lgStrSQL = lgStrSQL & "AND		ISSUE_REQ_NO	=	" & FilterVar(UCase(Trim(strdp_no)) ,"","S")  & " "
    
    ELSE                    'insert
        lgStrSQL = "INSERT INTO M_ISSUE_REQ_HDR_KO441  "
        lgStrSQL = lgStrSQL & " (                      "
        lgStrSQL = lgStrSQL & " 	PLANT_CD,          "
        lgStrSQL = lgStrSQL & " 	ISSUE_REQ_NO,      "
        lgStrSQL = lgStrSQL & " 	REQ_DT,            "
        lgStrSQL = lgStrSQL & " 	TRNS_TYPE,         "
        lgStrSQL = lgStrSQL & " 	ISSUE_TYPE,        "
        lgStrSQL = lgStrSQL & " 	MOV_TYPE,          "
        lgStrSQL = lgStrSQL & " 	EMP_NO,            "
        lgStrSQL = lgStrSQL & " 	DEPT_CD,           "
        lgStrSQL = lgStrSQL & " 	CONFIRM_FLAG,      "
        lgStrSQL = lgStrSQL & " 	REMARK,            "
        lgStrSQL = lgStrSQL & " 	INSRT_USER_ID,     "
        lgStrSQL = lgStrSQL & " 	INSRT_DT,          "
        lgStrSQL = lgStrSQL & " 	UPDT_USER_ID,      "
        lgStrSQL = lgStrSQL & " 	UPDT_DT,            "
        lgStrSQL = lgStrSQL & " 	LOC            "                '20080305::HANC
        lgStrSQL = lgStrSQL & " )                      "
        lgStrSQL = lgStrSQL & " VALUES                 "
        lgStrSQL = lgStrSQL & " (                      "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtPlantCd"))) ,"","S")         & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(strdp_no)) ,"","S")           & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtPoDt"))) ,"","S")         & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("rdoReleaseflg"))) ,"","S")      & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtPoTypeCd"))) ,"","S")        & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtPoTypeCd"))) ,"","S")        & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtEmp_no"))) ,"","S")       & ", "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtDept_cd"))) ,"","S")      & ", "
        lgStrSQL = lgStrSQL & " 	'N',               "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtRemark"))) ,"","S")      & ", "
        lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                          & "," 
        lgStrSQL = lgStrSQL & " 	GETDATE(),         "
        lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                          & "," 
        lgStrSQL = lgStrSQL & " 	GETDATE(),          "
        lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtLoc"))) ,"","S")      & " "      '20080305::hanc
        lgStrSQL = lgStrSQL & " )                      "
        
    END IF

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
 
'    If CheckSYSTEMError(Err,True) = True Then
'    	Call ServerMesgBox("stop"  , vbInformation, I_MKSCRIPT)          
'       lgErrorStatus    = "YES"
'       ObjectContext.SetAbort
'    End If
       
'20080221::HANC::BEGIN
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " parent.frm1.txtPoNo.value  = """ & UCase(Trim(strdp_no)) & """" & vbCr      
       Response.Write  " parent.frm1.txtPoNo1.value = """ & UCase(Trim(strdp_no)) & """" & vbCr      
       'Response.Write  " parent.frm1.txtPoNo1.value = 'ABCDE' " & vbCr      
       Response.Write  " </Script>                  " & vbCr
'20080221::HANC::END

End Sub	



'============================================================================================================
' Name : SubAuto_number
' Desc : 
'============================================================================================================
Sub SubAuto_number()

    Call SubCreateCommandObject(lgObjComm)
				
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    With lgObjComm
    
        .CommandText = "usp_z_auto_numbering_dvno"
        
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"   ,adInteger	,adParamReturnValue					)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@type"          ,adChar		,adParamInput  ,3  ,prefix			)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date"  ,adChar		,adParamInput  ,8  ,gl_dt			)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"        ,adChar		,adParamInput  ,13 ,gUsrId			)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@last_auto_no"  ,adChar		,adParamOutput ,18					)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@curr_dt"       ,adChar		,adParamInput  ,24 ,GetSvrDateTime	)

        lgObjComm.Execute ,, adExecuteNoRecords

    End With
   
    If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		if  IntRetCD <> 1 then		'정상실행일때 1 을 return 한다.

			Call DisplayMsgBox("", vbInformation, "자동채번생성오류입니다!!", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
		end if

		strdp_no = lgObjComm.Parameters("@last_auto_no").Value
'Call ServerMesgBox("PLEASE strdp_no : " & strdp_no  , vbInformation, I_MKSCRIPT)

	Else
		lgErrorStatus     = "YES"                                                         '☜: Set error status
		Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if
    
    Call SubCloseCommandObject(lgObjComm)
  
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

    lgStrSQL = "DELETE  FROM M_ISSUE_REQ_HDR_KO441"
    lgStrSQL = lgStrSQL & " WHERE PLANT_CD  = " &  FilterVar(UCase(Trim(Request("txtPlantCd"))) ,"","S") 
    lgStrSQL = lgStrSQL & " AND   ISSUE_REQ_NO = " & FilterVar(UCase(Trim(Request("txtPoNo"))) ,"","S") 
    

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If

'20080305::HANC::BEGIN
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "    Parent.DBSaveDelOk   " & vbCr      
       Response.Write  " </Script>                  " & vbCr
'20080305::HANC::END
    
End Sub




'============================================================================================================
' Name : SubRelease
' Desc : 발주확정 
'============================================================================================================
Sub SubRelease()

    Dim lgStrSQL
	Dim PM9G112
	Dim strMode,lgIntFlgMode
	Dim txtSpread
    Dim pvCB
	reDim IG1_import_group(0,2)
    Const M155_IG1_I1_select_char = 0 
    Const M155_IG1_I1_count = 1
    Const M155_IG1_I2_po_no = 2

	Dim prErrorPosition 
	Dim E3_m_pur_ord_hdr_po_no
	
    On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
        lgStrSQL =  "UPDATE  M_ISSUE_REQ_HDR_KO441 "
        lgStrSQL = lgStrSQL & "SET     CONFIRM_FLAG     =   " & FilterVar(UCase(Trim(Request("txtReleaseFlag"))) ,"","S")  & " "
        lgStrSQL = lgStrSQL & "WHERE	PLANT_CD		=	" & FilterVar(UCase(Trim(Request("txtPlantCd"))) ,"","S")  & " "
        lgStrSQL = lgStrSQL & "AND		ISSUE_REQ_NO	=	" & FilterVar(UCase(Trim(Request("txtPoNo"))) ,"","S")  & " "


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If



	Response.Write "<Script Language=vbscript>" 					& vbCr
	Response.Write "With parent"									& vbCr	
    Response.Write  " Parent.DBSaveOk            " & vbCr
	Response.Write "End With"   & vbCr
	Response.Write "</Script>" & vbCr

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


