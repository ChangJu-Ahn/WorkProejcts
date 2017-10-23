<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Dim lgSeq
	Dim lgQty
	Dim user_id
	Dim prefix
	Dim gl_dt
	Dim strdp_no, IntRetCD
	
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    user_id = Request("txtInsrtUserId") 
    gl_dt   = Request("txtcurr_dt") 
    prefix  = "DP"
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			 Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "PREFIXGEN"															 '☜: store proecedure
             Call SubCreateCommandObject(lgObjComm)
             Call SubBizBatch()
             Call SubCloseCommandObject(lgObjComm)
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizBatch()
' Desc : Query Data from Db
'============================================================================================================
Sub  SubBizBatch()

    Dim strMsg_cd 
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim strYMD , strBP , strIP

    strYMD = REPLACE(lgKeyStream(0),"-","")

    If lgKeyStream(1) <> "" Then 
       strBP  = lgKeyStream(1)
    Else 
       strBP  = "NULL"
    End If

    strIP  = lgKeyStream(2)

    With lgObjComm
    
        .CommandText = "usp_a_tempgl_no_auto_gen"
        .CommandType = adCmdStoredProc       

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"  ,adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@type"         ,adVarChar,adParamInput,  3, strtype)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date" ,adVarChar,adParamInput,  8, strdt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"       ,adVarChar,adParamInput,  13, strid)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@last_auto_no" ,adVarChar,adParamoutput, 18)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@curr_dt"      ,adVarChar,adParamInput , 24, FilterVar(GetSvrDateTime,"''","S"))
	    
	    .Execute ,, adExecuteNoRecords
    
    End With
    
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        If  Cdbl(IntRetCD) < 0 Then
			'strMsg_cd = lgObjComm.Parameters("@MSG_NO").Value

            'Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
            
            IntRetCD = -1
            'Exit Sub
        Else
            IntRetCD = 1
        End if
    Else           
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
    If Cdbl(IntRetCD) > 0 Then
'       Response.Write  " <Script Language=vbscript> " & vbCr
'       Response.Write  "    Parent.ExeReflectOk     " & vbCr      
'       Response.Write  " </Script>                  " & vbCr
        CALL SubBizQueryMulti()
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "    Parent.ExeReflectNo     " & vbCr      
       Response.Write  " </Script>                  " & vbCr
    End If
          
End Sub	


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Call SubMakeSQLStatements("S")                                                  '☜ : Make sql statements ,SR : Single Read
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'R(Read) X(CursorType) X(LockType) 
       If lgPrevNext = "Q" Then

          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found. 
          
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the starting data. 
          lgPrevNext = "Q"
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the ending data.
          lgPrevNext = "Q"
          Call SubBizQuery()
       End If
    Else
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
       Response.Write  " <Script Language=vbscript>" & vbCr
       Response.Write  " With Parent               " & vbCr
       
       Response.Write  "       .Frm1.txtDlvyNo2.Value   = """ & ConvSPChars(lgObjRs("DLVY_NO"))          & """" & vbCr
       Response.Write  "       .Frm1.txtDlvyTime.Value	= """ & ConvSPChars(lgObjRs("DLVY_TIME"))        & """" & vbCr
       Response.Write  "       .Frm1.txtTransCo.Value	= """ & ConvSPChars(lgObjRs("TRANS_CO"))         & """" & vbCr
       Response.Write  "       .Frm1.txtVehicleNo.Value	= """ & ConvSPChars(lgObjRs("VEHICLE_NO"))       & """" & vbCr
       Response.Write  "       .Frm1.txtDriver.Value	= """ & ConvSPChars(lgObjRs("DRIVER"))           & """" & vbCr
       Response.Write  "       .Frm1.txtTelNo1.Value	= """ & ConvSPChars(lgObjRs("TEL_NO1"))          & """" & vbCr
       Response.Write  "       .Frm1.txtTelNo2.Value	= """ & ConvSPChars(lgObjRs("TEL_NO2"))          & """" & vbCr
       Response.Write  "       .Frm1.txtDlvyPlace.Value	= """ & ConvSPChars(lgObjRs("DLVY_PLACE"))       & """" & vbCr
       Response.Write  "       .Frm1.txtREMARK.Value	= """ & ConvSPChars(lgObjRs("REMARK"))           & """" & vbCr
       
       'Response.Write  "       .DBQueryOk " & vbCr
       Response.Write  " End With         " & vbCr
       Response.Write  " </Script>        " & vbCr
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
    End If

    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
    
    Call SubBizQueryCond()
    If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	End If
End Sub    

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    Dim lgIntFlgMode
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                '☜: Read Operayion Mode (CREATE, UPDATE)

   
    if lgIntFlgMode = OPMD_CMODE then
       Call SubAuto_number()
    end if
   
    Select Case lgIntFlgMode
        Case  OPMD_CMODE  Call SubBizSaveSingleCreate()                            '☜ : Create
        Case  OPMD_UMODE  : Call SubBizSaveSingleUpdate()                            '☜ : Update
    End Select
	
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

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@type"          ,adChar,adParamInput, 3,  prefix)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date"  ,adChar,adParamInput, 8, gl_dt)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"        ,adChar,adParamInput, 13,  gUsrId)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@last_auto_no"  ,adChar,adParamOutput, 18)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@curr_dt"       ,adChar,adParamInput, 24,  GetSvrDateTime)

        lgObjComm.Execute ,, adExecuteNoRecords

    End With
   
    If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		if  IntRetCD <> 1 then		'정상실행일때 1 을 return 한다.

			Call DisplayMsgBox("", vbInformation, "자동채번생성오류입니다!!", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
		end if

		strdp_no = lgObjComm.Parameters("@last_auto_no").Value

	Else
		lgErrorStatus     = "YES"                                                         '☜: Set error status
		Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if
    
    Call SubCloseCommandObject(lgObjComm)
        
End Sub	    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL =            "" 
	lgStrSQL = lgStrSQL & " DELETE FROM M_SCM_DLVY_PUR_RCPT "
	lgStrSQL = lgStrSQL & "  WHERE DLVY_NO = " & FilterVar(lgKeyStream(1) & "" ,"''", "S") 
	lgStrSQL = lgStrSQL & "    AND BP_CD   = " & FilterVar(lgKeyStream(0) & "" ,"''", "S") 

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    lgStrSQL =            "" 
	lgStrSQL = lgStrSQL & " UPDATE M_SCM_FIRM_PUR_RCPT SET "
	lgStrSQL = lgStrSQL & "        DLVY_NO     = NULL , "
	lgStrSQL = lgStrSQL & "        DLVY_SEQ_NO = 0 "
	lgStrSQL = lgStrSQL & "  WHERE DLVY_NO = " & FilterVar(lgKeyStream(1) & "" ,"''", "S") 
	
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    
    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.DbDeleteOk    " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If   

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim lgStrSQL


    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

	lgStrSQL =            " INSERT INTO M_SCM_DLVY_PUR_RCPT "
	lgStrSQL = lgStrSQL & " (BP_CD , "
	lgStrSQL = lgStrSQL & " DLVY_NO, "
	lgStrSQL = lgStrSQL & " DLVY_TIME, "
	lgStrSQL = lgStrSQL & " TRANS_CO, "
	lgStrSQL = lgStrSQL & " VEHICLE_NO, "
	lgStrSQL = lgStrSQL & " DRIVER, "
	lgStrSQL = lgStrSQL & " TEL_NO1, "
	lgStrSQL = lgStrSQL & " TEL_NO2, "
	lgStrSQL = lgStrSQL & " DLVY_PLACE , "
	lgStrSQL = lgStrSQL & " REMARK, "
	lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
	lgStrSQL = lgStrSQL & " INSRT_DT, "
	lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
	lgStrSQL = lgStrSQL & " UPDT_DT) "
	lgStrSQL = lgStrSQL & " VALUES( "
	lgStrSQL = lgStrSQL & FilterVar(Request("txtbpcd")      ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(strdp_no    ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtdlvytime")  ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txttransco")      ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtvehicleno") ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtdriver")    ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txttelno1")    ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txttelno2")    ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtdlvyplace") ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtremark")    ,"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                     & "," 
	lgStrSQL = lgStrSQL & " GETDATE(), "
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                     & "," 
	lgStrSQL = lgStrSQL & " GETDATE()) "
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    End If   
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    Dim lgStrSQL
    Dim tmpDate

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    tmpDate = FilterVar(UniConvDateAToB(Trim(Request("fpdtCloseDt")),gDateFormatYYYYMM,gServerDateFormat),null,"S")
    
    lgStrSQL = ""
	lgStrSQL = lgStrSQL & " UPDATE M_SCM_DLVY_PUR_RCPT SET "
	lgStrSQL = lgStrSQL & " 	DLVY_TIME	= " &  FilterVar(Request("txtdlvytime") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	TRANS_CO	= " &  FilterVar(Request("txttransco") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	VEHICLE_NO	= " &  FilterVar(Request("txtvehicleno") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	DRIVER		= " &  FilterVar(Request("txtdriver") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	TEL_NO1		= " &  FilterVar(Request("txttelno1") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	TEL_NO2		= " &  FilterVar(Request("txttelno2") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	DLVY_PLACE	= " &  FilterVar(Request("txtdlvyplace") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	REMARK		= " &  FilterVar(Request("txtremark") ,"''","S")  & ","
	lgStrSQL = lgStrSQL & " 	UPDT_USER_ID	= " &  FilterVar(gUsrId,"","S")  & ","
	lgStrSQL = lgStrSQL & " 	UPDT_DT		= GETDATE() "
	lgStrSQL = lgStrSQL & "  WHERE BP_CD   		= " &  FilterVar(Request("txtbpcd") ,"''","S")  
	lgStrSQL = lgStrSQL & "    AND DLVY_NO 		= " &  FilterVar(Request("txtdlvyno2") ,"''","S")  
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       'Response.Write  " <Script Language=vbscript> " & vbCr
       'Response.Write  "       Parent.DBSaveOk      " & vbCr
       'Response.Write  " </Script>                  " & vbCr
    End If   

End Sub

'============================================================================================================
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear

	If lgKeyStream(0) <> "" AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM003", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next
    Err.Clear

    strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
    '    Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
    '    Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("D_BP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))
            
            If ConvSPChars(lgObjRs("RET_FLG")) = "N" Then 
				lgstrData = lgstrData & Chr(11) & "정상"
            Else
				lgstrData = lgstrData & Chr(11) & "반품"
            End If
            
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

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)
		
		Select Case arrColVal(0)
            Case "U"                            '☜: Create 
				Call SubBizQuerySeq(arrColVal)
				Call SubBizSaveMultiUpdate(arrColVal)
			Case "D"
				Call SubBizSaveMultiDelete(arrColVal)
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuerySeq(arrColVal)
    On Error Resume Next
    Err.Clear

	lgStrSQL =            " SELECT isNULL(MAX(DLVY_SEQ_NO),0) DLVY_SEQ_NO "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(nolock) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO = " & FilterVar(Trim(UCase(arrColVal(2))),"","S") 
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		If lgObjRs("DLVY_SEQ_NO") = "" OR lgObjRs("DLVY_SEQ_NO") = 0 Then
			lgSeq = 1
		Else
			lgSeq = lgObjRs("DLVY_SEQ_NO") + 1
		End If
    End If
    
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL = "INSERT INTO M_SCM_FIRM_PUR_RCPT ( "
    lgStrSQL = lgStrSQL & vbCrLf & " PO_NO	, "
    lgStrSQL = lgStrSQL & vbCrLf & " PO_SEQ_NO	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " SPLIT_SEQ_NO	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " PLAN_DVRY_DT	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " PLAN_DVRY_QTY	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " M_TYPE	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " D_BP_CD	, "
    lgStrSQL = lgStrSQL & vbCrLf & " LOT_NO	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " CONFIRM_QTY	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " INSRT_USER_ID	, "
    lgStrSQL = lgStrSQL & vbCrLf & " INSRT_DT		, "  
    lgStrSQL = lgStrSQL & vbCrLf & " UPDT_USER_ID	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " UPDT_DT			) "    
    lgStrSQL = lgStrSQL & vbCrLf & " VALUES			( "
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(2))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(3))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & lgSeq  & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(4))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(5))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(6))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(7))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & "'*' ,"
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(5))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			& "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			& ")"     
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL =                     " UPDATE M_SCM_FIRM_PUR_RCPT SET "
	'lgStrSQL = lgStrSQL & vbCrLf & "        DLVY_NO       = " & FilterVar(Request("txtdlvyno2") ,"''","S") & ","
	lgStrSQL = lgStrSQL & vbCrLf & "        DLVY_NO       = " & FilterVar(strdp_no ,"''","S") & ","
	lgStrSQL = lgStrSQL & vbCrLf & "        DLVY_SEQ_NO   = " & lgSeq    
	lgStrSQL = lgStrSQL & vbCrLf & "  WHERE PO_NO         = " & FilterVar(Trim(UCase(arrColVal(02))),"","S")  
	lgStrSQL = lgStrSQL & vbCrLf & "    AND PO_SEQ_NO     = " & FilterVar(Trim(UCase(arrColVal(03))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND SPLIT_SEQ_NO = "  & FilterVar(Trim(UCase(arrColVal(12))),"","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL =                     " UPDATE M_SCM_FIRM_PUR_RCPT SET "
	lgStrSQL = lgStrSQL & vbCrLf & "        DLVY_NO       = NULL ,"
	lgStrSQL = lgStrSQL & vbCrLf & "        DLVY_SEQ_NO   = 0 "     
	lgStrSQL = lgStrSQL & vbCrLf & "  WHERE PO_NO         = " & FilterVar(Trim(UCase(arrColVal(02))),"","S")  
	lgStrSQL = lgStrSQL & vbCrLf & "    AND PO_SEQ_NO     = " & FilterVar(Trim(UCase(arrColVal(03))),"","S")
	lgStrSQL = lgStrSQL & vbCrLf & "    AND SPLIT_SEQ_NO = "  & FilterVar(Trim(UCase(arrColVal(04))),"","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
					lgStrSQL =            "	SELECT TOP " & iSelCount  & " A.PO_NO , A.PO_SEQ_NO , A.SPLIT_SEQ_NO , B.ITEM_CD , C.ITEM_NM , C.SPEC , "
					lgStrSQL = lgStrSQL & "	       C.BASIC_UNIT , A.PLAN_DVRY_DT , A.PLAN_DVRY_QTY , A.D_BP_CD ,  D.SL_NM , E.RET_FLG "
					lgStrSQL = lgStrSQL & "	  FROM M_SCM_FIRM_PUR_RCPT A, m_pur_ord_dtl B, B_ITEM C , B_STORAGE_LOCATION D , M_PUR_ORD_HDR E "
					lgStrSQL = lgStrSQL & "	 WHERE A.PO_NO     = B.PO_NO  "
					lgStrSQL = lgStrSQL & "	   AND A.PO_SEQ_NO = B.PO_SEQ_NO  "
					lgStrSQL = lgStrSQL & "	   AND B.ITEM_CD = C.ITEM_CD "
					lgStrSQL = lgStrSQL & "	   AND A.D_BP_CD = D.SL_CD "
					lgStrSQL = lgStrSQL & "	   AND E.PO_NO   = B.PO_NO  "

					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "  AND E.BP_CD = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
					End If				
	
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  AND A.DLVY_NO = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
					End If
					
					
           End Select             

        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
           End Select 
           
		Case "S"
			
			lgStrSQL =            " SELECT * " 
			lgStrSQL = lgStrSQL & "   FROM M_SCM_DLVY_PUR_RCPT "
			lgStrSQL = lgStrSQL & "  WHERE BP_CD   =  " & FilterVar(lgKeyStream(0),"''", "S")
			lgStrSQL = lgStrSQL & "    AND DLVY_NO =  " & FilterVar(lgKeyStream(1),"''", "S")
			
    End Select
    
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
        
			If CheckSYSTEMError(pErr,True) = True Then
			   ObjectContext.SetAbort
			   Call SetErrorStatus
			Else
			   If CheckSQLError(pConn,True) = True Then
			      ObjectContext.SetAbort
			      Call SetErrorStatus
			   End If
			End If
			
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk("<%=lgStrPrevKey%>")
	         End with
	      Else
				Parent.DBQueryNotOk()
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             parent.frm1.txtDlvyNo.value = "<%=ConvSPChars(strdp_no)%>"
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               