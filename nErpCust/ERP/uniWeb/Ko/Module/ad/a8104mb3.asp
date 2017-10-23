<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%
         
    Dim lgOpModeCRUD
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	lgOpModeCRUD      = Request("txtMode")                                          '¢Ð: Read Operation Mode (CRUD)
	Select Case lgOpModeCRUD    
'		Case CStr(UID_M0001)                                                         '¢Ð: Query
'			Call SubBizQuery()
'			Call SubBizQueryMulti()
'		Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
'			Call SubBizSave()
'			Call SubBizSaveMulti()
		Case CStr(UID_M0003)                                                         '¢Ð: Delete
			Call SubBizDelete()
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    Const A359_I3_allc_no = 0    '  View Name : import a_allc_rcpt
    Const A359_I3_allc_dt = 1
    Const A359_I3_allc_type = 2
    Const A359_I3_ref_no = 3
    Const A359_I3_allc_amt = 4
    Const A359_I3_allc_loc_amt = 5
    Const A359_I3_dc_amt = 6
    Const A359_I3_dc_loc_amt = 7
    Const A359_I3_insrt_user_id = 8
    Const A359_I3_insrt_dt = 9
    Const A359_I3_updt_user_id = 10
    Const A359_I3_updt_dt = 11
	
	Dim objPADG020
	Dim strAllcNo
	Dim I3_a_allc_rcpt
	Dim iCommandSent

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
	
	strAllcNo = Request("txtAllcNo")
	
	If Trim(strAllcNo) <> "" Then
		Redim I3_a_allc_rcpt(A359_I3_updt_dt)	
		I3_a_allc_rcpt(A359_I3_allc_no) = strAllcNo
	End If
	
	iCommandSent = "DELETE"
	
	Set objPADG020 = CreateObject("PADG020.cAMntAllcRcHqSvr")
	
    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
		
	Call objPADG020.A_MAINT_ALLC_RCPT_HQ_SVR (gStrGloBalCollection, iCommandSent, , , ,I3_a_allc_rcpt)
										 
	If CheckSYSTEMError(Err, True) = True Then
       Set objPADG020 = Nothing
		Exit Sub
    End If    										 
        
	Set objPADG020 = nothing

    
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr		
	Response.Write " 	.DbDeleteOK				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr 	
	
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear          
	                                                             
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr
	Response.Write " 	.DbQueryOk				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr 
	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	
	Response.Write "<Script Language=vbscript>	" & vbcr	
	Response.Write " With Parent				" & vbCr	
	Response.Write " 	.DbSaveOk				" & vbCr	
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr
	
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

%>
