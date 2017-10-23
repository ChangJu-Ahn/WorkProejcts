
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Open A/P Confirm
'*  3. Program ID           : a5103mb2
'*  4. Program Name         :  
'*  5. Program Desc         :  
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Chang Goo,Kang
'* 10. Modifier (Last)      : Ahn Hae Jin
'* 11. Comment              :
'*
'**********************************************************************************************
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<% 
	Call LoadBasisGlobalInf() 
%>
<%
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call ggHideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             'Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
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



End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	
	Dim PAGG015_cACnfmTmpGlSvr
	Dim iStrWkFg
    
	Dim iCommandSent
	Dim I1_from_temp_gl_dt
	Dim I2_to_temp_gl_dt
	Dim I3_b_acct_dept
	Dim I4_gl_input_type
	Dim I5_issued_dt
	Dim l6_from_temp_gl_no
	Dim l7_to_temp_gl_no
	Dim IG1_import_grp_temp_gl
	Dim iErrorPosition
    
    Redim I3_b_acct_dept(1)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    
    iStrWkFg = Request("htxtWorkFg")
   
   	If UCase(iStrWkFg)	  = "CONF"    then		
		iCommandSent   = "CONF"			
	Elseif UCase(strWkfg) = "UNCONF"  then		
		iCommandSent   = "UNCONF"
	End if
    
        
    I1_from_temp_gl_dt	= UNIConvDate(Request("txtFromReqDt"))    
    I2_to_temp_gl_dt	= UNIConvDate(Request("txtToReqDt"))     
    I3_b_acct_dept(0)	= Trim(Request("hOrgChangeId"))				'GetGlobalInf("gChangeOrgId")
    I3_b_acct_dept(1)	= Trim(Request("txtDeptCd"))
    I4_gl_input_type	= Trim(Request("txtGlInputType"))
    I5_issued_dt		= UNIConvDate(Trim(Request("GIDate")))
    l6_from_temp_gl_no	= Request("txtTempGlNoFr")
    l7_to_temp_gl_no	= Request("txtTempGlNoTo")
	
    Set PAGG015_cACnfmTmpGlSvr = Server.CreateObject("PAGG015.cACnfmTmpGlSvr")
    
    If CheckSYSTEMError(Err, True) = True Then					       
       Call SetErrorStatus
       Exit Sub
    End If    
													
	Call PAGG015_cACnfmTmpGlSvr.A_CONFIRM_TEMP_GL_SVR(gStrGlobalCollection, _
													iCommandSent, _
													I1_from_temp_gl_dt, _
													I2_to_temp_gl_dt, _
													I3_b_acct_dept, _
													I4_gl_input_type, _
													I5_issued_dt, _
													l6_from_temp_gl_no, _
													l7_to_temp_gl_no, _
													IG1_import_grp_temp_gl,_
													iErrorPosition, _
													gDsnNo)	
													
	If CheckSYSTEMError2(Err, True,iErrorPosition & "","","","","") = True Then
		Set PAGG015_cACnfmTmpGlSvr = Nothing
		Call SetErrorStatus
		'Exit Sub
	End If

    
    Set PAGG015_cACnfmTmpGlSvr = Nothing
    Response.Write " <Script Language=vbscript>									" & vbCr	
	Response.Write " With Parent												" & vbCr
	Response.Write "	If """ & lgErrorStatus & """ = ""NO"" then				" & vbCr
	Response.Write "		.InitSpreadSheet									" & vbCr	
	Response.Write "		.InitVariables										" & vbCr	
	Response.Write "		If """ & UCase(iStrWkFg) & """ = ""CONF"" Then		" & vbCr
	Response.Write "			.frm1.cboConfFg.value = ""C""					" & vbCr
	Response.Write "		ElseIf """ & UCase(iStrWkFg) & """ = ""UNCONF"" Then" & vbCr
	Response.Write "			.frm1.cboConfFg.value = ""U""					" & vbCr
	Response.Write "		End If												" & vbCr
	Response.Write "		.DbQuery											" & vbCr
	Response.Write "		.cboConfFg_onchange									" & vbCr
	Response.Write "	Else													" & vbCr
	Response.Write "		Call .LayerShowHide(0)								" & vbCr
	Response.Write "		.Frm1.vspdData.Focus								" & vbCr
	Response.Write "	End If													" & vbCr	
	Response.Write " End With													" & vbCr
	Response.Write " </Script>													" & vbCr  

	
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MU"
    End Select
End Sub

%>



