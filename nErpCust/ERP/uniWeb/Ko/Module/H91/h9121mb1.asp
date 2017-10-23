<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    Dim lgYearAreaCd
    Dim lgYearAreaBodyCd

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","H","NOCOOKIE","MB")   
    
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    
    lgYearAreaCd		= FilterVar(UCase(Request("txtYearAreaCd")), "''", "S")
    lgYearAreaBodyCd	= FilterVar(UCase(Request("txtYearAreaCd_Body")), "''", "S")

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection


    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call SubMakeSQLStatements("R", lgYearAreaCd, "")                                   '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If

    Else%>
<Script language="VBScript">
	With Parent.frm1 
		.txtYearAreaNm.value = "<%=ConvSPChars(lgObjRs("YEAR_AREA_NM"))%>"
		.txtYearAreaCd_Body.value = "<%=ConvSPChars(lgObjRs("YEAR_AREA_CD"))%>"
		.txtYearAreaNm_Body.value = "<%=ConvSPChars(lgObjRs("YEAR_AREA_NM"))%>"
		.txtOwnRgstNo.value = "<%=ConvSPChars(lgObjRs("OWN_RGST_NO"))%>"
		.txtCoOwnRgstNo.value = "<%=ConvSPChars(lgObjRs("CO_OWN_RGST_NO"))%>"
		.txtRepreNm.value = "<%=ConvSPChars(lgObjRs("REPRE_NM"))%>"
		.txtTaxOffice.value = "<%=ConvSPChars(lgObjRs("tax_biz_cd"))%>"
		.txtTaxOfficeNm.value = "<%=ConvSPChars(lgObjRs("tax_biz_nm"))%>"
		.txtTelNo.value = "<%=ConvSPChars(lgObjRs("TEL_NO"))%>"  
		.txtAddr.value = "<%=ConvSPChars(lgObjRs("ADDR"))%>" 
		.txtWorkerNm.value = "<%=ConvSPChars(lgObjRs("WORKER_NAME"))%>" 
		.txtWorkerDeptNm.value = "<%=ConvSPChars(lgObjRs("WORKER_DEPT_NM"))%>" 
		.txtWorkerTel.value = "<%=ConvSPChars(lgObjRs("WORKER_TEL"))%>" 	
		.txtHometaxID.value = "<%=ConvSPChars(lgObjRs("HOMETAX_ID"))%>" 								
	End With                                                       

</Script>  	
    <%End If
    Call SubCloseRs(lgObjRs)    
                                                    '¢Ð : Release RecordSSet
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Dim lgIntFlgMode

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '¢Ð : Update
              Call SubBizSaveSingleUpdate()
        Case  OPMD_DMODE                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL =  " DELETE HFA100T"
    lgStrSQL = lgStrSQL & "  WHERE YEAR_AREA_CD = " & lgYearAreaBodyCd

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
	
	Call SubCloseRs(lgObjRs)                                                    '¢Ð : Release RecordSSet

End Sub
'============================================================================================================
' Name : SubBizUpdateAfterDelete
' Desc : Update After Delete DB data When Delete Work is Succeed
'============================================================================================================

'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	lgStrSQL =  		  " INSERT INTO HFA100T "
	lgStrSQL = lgStrSQL & "(YEAR_AREA_CD ,"     
	lgStrSQL = lgStrSQL & "YEAR_AREA_NM     , "
	lgStrSQL = lgStrSQL & "OWN_RGST_NO      , "
	lgStrSQL = lgStrSQL & "CO_OWN_RGST_NO   , "
	lgStrSQL = lgStrSQL & "REPRE_NM         , "
	lgStrSQL = lgStrSQL & "TEL_NO           , "
	lgStrSQL = lgStrSQL & "ADDR             , "
	lgStrSQL = lgStrSQL & "tax_biz_cd       , "
	lgStrSQL = lgStrSQL & "tax_biz_nm       , "
	lgStrSQL = lgStrSQL & "WORKER_NAME		, "	'2004
	lgStrSQL = lgStrSQL & "WORKER_DEPT_NM	, " '2004
	lgStrSQL = lgStrSQL & "WORKER_TEL	, "		'2004	
	lgStrSQL = lgStrSQL & "HOMETAX_ID	, "		'2004		
	lgStrSQL = lgStrSQL & "ISRT_EMP_NO      , "
	lgStrSQL = lgStrSQL & "ISRT_DT          , "
	lgStrSQL = lgStrSQL & "UPDT_EMP_NO      , "
	lgStrSQL = lgStrSQL & "UPDT_DT            "
	lgStrSQL = lgStrSQL & " )"  
    lgStrSQL = lgStrSQL & " VALUES (" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtYearAreaCd_Body"))), "''","S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Trim(Request("txtYearAreaNm_Body"))), "''","S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtOwnRgstNo")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtCoOwnRgstNo")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtRepreNm")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtTelNo")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtAddr")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtTaxOffice")), "''","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtTaxOfficeNm")), "''","S")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtWorkerNm")), "''","S")     & ","		'2004
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtWorkerDeptNm")), "''","S")     & ","	'2004
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtWorkerTel")), "''","S")     & ","		'2004
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtHometaxID")), "''","S")     & ","		'2004
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''","S")   & "," 
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''","S")  & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''","S")   & "," 
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''","S")
    lgStrSQL = lgStrSQL & ")"

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	lgStrSQL = "UPDATE  HFA100T"
	lgStrSQL = lgStrSQL & " SET " 
	lgStrSQL = lgStrSQL & " YEAR_AREA_NM = " & FilterVar(Trim(Request("txtYearAreaNm_Body")), "''","S")  & ","
	lgStrSQL = lgStrSQL & " OWN_RGST_NO = " & FilterVar(Request("txtOwnRgstNo"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " CO_OWN_RGST_NO = " & FilterVar(Request("txtCoOwnRgstNo"), "''", "S") & ","			 
	lgStrSQL = lgStrSQL & " REPRE_NM = " & FilterVar(Request("txtRepreNm"), "''", "S") & ","	
	lgStrSQL = lgStrSQL & " TEL_NO = " & FilterVar(Request("txtTelNo"), "''", "S") & ","    
	lgStrSQL = lgStrSQL & " ADDR = " & FilterVar(Request("txtAddr"), "''", "S") & ","    
	lgStrSQL = lgStrSQL & " tax_biz_cd = " & FilterVar(Request("txtTaxOffice"), "''", "S") & ","    
    lgStrSQL = lgStrSQL & " tax_biz_nm = " & FilterVar(Request("txtTaxOfficeNm"), "''", "S") & "," 
	lgStrSQL = lgStrSQL & " WORKER_NAME = " & FilterVar(Request("txtWorkerNm"), "''", "S") & ","    
	lgStrSQL = lgStrSQL & " WORKER_DEPT_NM = " & FilterVar(Request("txtWorkerDeptNm"), "''", "S") & ","    
	lgStrSQL = lgStrSQL & " WORKER_TEL = " & FilterVar(Request("txtWorkerTel"), "''", "S") & ","    
	lgStrSQL = lgStrSQL & " HOMETAX_ID = " & FilterVar(Request("txtHometaxID"), "''", "S") & ","    	
	lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''","S") & ","   
	lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime, "''", "S")
	lgStrSQL = lgStrSQL & " WHERE YEAR_AREA_CD = " & lgYearAreaCd
  
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
	
	Call SubCloseRs(lgObjRs)                                                    '¢Ð : Release RecordSSet

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pStrMode,pCode, pCode1)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	Dim pMode

	pMode = left(pStrMode,1)

    Select Case pMode 
      Case "R"
            lgStrSQL =  " 		select  YEAR_AREA_CD,"
			lgStrSQL = lgStrSQL & " 	YEAR_AREA_NM,"
			lgStrSQL = lgStrSQL & " 	CO_OWN_RGST_NO,"
			lgStrSQL = lgStrSQL & " 	tax_biz_cd,"
			lgStrSQL = lgStrSQL & " 	tax_biz_nm,"
			lgStrSQL = lgStrSQL & " 	OWN_RGST_NO,"
			lgStrSQL = lgStrSQL & " 	REPRE_NM,"
			lgStrSQL = lgStrSQL & " 	TEL_NO,ADDR,"
			lgStrSQL = lgStrSQL & " 	WORKER_NAME,WORKER_DEPT_NM,WORKER_TEL, HOMETAX_ID" '2004
			lgStrSQL = lgStrSQL & " from HFA100T "
			lgStrSQL = lgStrSQL & " WHERE YEAR_AREA_CD = " & pCode 
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "SC"		'//insert
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
					End If
                 End If

        Case "SD"		'//delete
				If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)    
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)   
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SU"		'//update
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990024", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990024", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select

End Sub


%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"		
          If Trim("<%=lgErrorStatus%>") = "NO" Then			
             Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
