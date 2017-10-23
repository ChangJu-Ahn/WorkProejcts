
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5123mb2
'*  4. Program Name         : 회계전표일괄생성 
'*  5. Program Desc         : 각 모쥴에서 생성한 자료를 토대로 일괄적으로 전표처리.
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/09/26 : ..........
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


<% 
	Call LoadBasisGlobalInf() 
%>
<%
     
    Dim lgOpModeCRUD
    Dim lgstrWkfg
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgStrWkfg			= Request("htxtWorkFg")
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
             Call SubBizDeleteMulti()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
		
	Const A377_I2_from_a_batch_gl_dt			= 0
	Const A377_I2_from_a_batch_gl_input_type	= 1

	Const A377_IG1_a_batch_batch_no				= 0
	Const A377_IG1_a_batch_auto_trans_fg		= 1
	
	Dim PAGG115_cAMngBtchToGlSvr
	Dim iCommandSent
	Dim I1_b_biz_area
	Dim I2_from_a_batch
	Dim I3_to_a_batch
	Dim IG1_a_batch
	Dim iErrorPosition
	
	ReDim I2_from_a_batch(A377_I2_from_a_batch_gl_input_type)
	I1_b_biz_area											= Trim(Request("txtBizCd"))
	I2_from_a_batch(A377_I2_from_a_batch_gl_dt)				= UNIConvDate(Request("txtFromReqDt"))
	I2_from_a_batch(A377_I2_from_a_batch_gl_input_type)		= Trim(Request("txtGlInputType"))
	I3_to_a_batch											= UNIConvDate(Request("txtToReqDt"))
	
	
		
	If UCase(lgStrWkfg) = "CONF"  then		
		iCommandSent    = "CONF"			
	Else
		iCommandSent    = "UNCONF"
	End if
		
	
	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------	
	
	
	Set PAGG115_cAMngBtchToGlSvr = Server.CreateObject("PAGG115.cAMngBtchToGlSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
	
	Call PAGG115_cAMngBtchToGlSvr.A_MANAGE_BATCH_TO_GL_SVR(gStrGlobalCollection, _
															iCommandSent, _
															I1_b_biz_area, _
															I2_from_a_batch, _
															I3_to_a_batch, _
															IG1_a_batch, _
															iErrorPosition) 
	
'	if err.number <> 0 then
'		Response.Write "xx  "
'		Response.Write err.description & " :: " & err.source
'		Set PAGG005_cAMngTmpGlSvr  = Nothing
'		Response.End
'	end if
	
	Response.Write I2_from_a_batch(A377_I2_from_a_batch_gl_input_type)	
	If CheckSYSTEMError2(Err, True,iErrorPosition & "","","","","") = True Then
		Set PAGG115_cAListBtchSvr = Nothing
		Call SetErrorStatus
		Exit Sub
	End If
	
    Set PAGG115_cAMngBtchToGlSvr  = Nothing

	Response.Write " <Script Language=vbscript>									" & vbCr	
	Response.Write " With Parent												" & vbCr
	Response.Write "	.InitSpreadSheet										" & vbCr
	Response.Write "	.InitVariables											" & vbCr	
	Response.Write "	If """ & UCase(lgStrWkfg) & """  = ""CONF"" Then		" & vbCr
	Response.Write "		.frm1.cboConfFg.value = ""C""						" & vbCr	
	Response.Write "	ElseIf """ & UCase(lgStrWkfg) & """ = ""UNCONF"" Then	" & vbCr
	Response.Write "		.frm1.cboConfFg.value = ""U""						" & vbCr
	Response.Write "	End If													" & vbCr
	Response.Write "	.DbQuery												" & vbCr
	Response.Write "	.cboConfFg_OnChange										" & vbCr	
	Response.Write " End With													" & vbCr
	Response.Write " </Script>													" & vbCr 

End Sub    



'============================================================================================================
' Name : SubBizDeleteMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDeleteMulti()

	
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
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
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
End Sub

%>


