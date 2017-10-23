<%@ LANGUAGE=VBSCript%>
<%  
Option Explicit  

%>

<!-- #Include file="../../inc/IncServer.asp" -->

<%
    Dim lgOpModeCRUD
    
    On Error Resume Next															'¢Ð: Protect system from crashing
    Err.Clear																		'¢Ð: Clear Error status

    Call HideStatusWnd																'¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")											'¢Ð: Read Operation Mode (CRUD)
    Call SubBizDelete()

'    Select Case lgOpModeCRUD
'        Case CStr(UID_M0001)														'¢Ð: Query
'           'Call SubBizQuery()														'¢Ð: Single --> Query
'             Call SubBizQueryMulti()												'¢Ð: Multi  --> Query
'        Case CStr(UID_M0002)														'¢Ð: Save,Update
'            'Call SubBizSave()														'¢Ð: Single --> Save,Update
'             Call SubBizSaveMulti()													'¢Ð: Multi  --> Save,Update,Delete
'        Case CStr(UID_M0003)														'¢Ð: Delete
'            													    '¢Ð: Single --> Delete
'    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    'On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPAPG020
    Dim iCommandSent
    Dim I9_batch_paym_no
    Dim I5_a_allc_paym

    Const A363_I5_paym_type = 8
	Const A363_I5_prpaym_no = 18
	
	redim I5_a_allc_paym(A363_I5_prpaym_no)
	    
    iCommandSent = "DELETE"

    I9_batch_paym_no = Trim(Request("txtBatchAllcNo"))
	I5_a_allc_paym(A363_I5_paym_type) = Trim(UCase(Request("txtInputType")))    
    
    Set iPAPG020 = Server.CreateObject ("PAPG020.cAMntPayAllcSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
	
	Response.Write I5_a_allc_paym(A363_I5_paym_type)
	Response.END 
    Call iPAPG020.A_MAINT_BATCH_PAYM_ALLC_SVR(gStrGlobalCollection,iCommandSent,,, , ,I5_a_allc_paym , , , , , I9_batch_paym_no)	
	
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG020 = Nothing
		Exit Sub
	End If

    Set iPAPG020 = Nothing
    
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write " parent.DbDeleteOk()  " & vbCr  
    Response.Write "</Script>"  
        
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>