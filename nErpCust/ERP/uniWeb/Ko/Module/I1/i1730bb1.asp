<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->

<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
     
    Call LoadBasisGlobalInf() 
    'Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    
	On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)


	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizClose()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizCancel()
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizClose()
	Dim IntRetCD
	Dim strMsg_cd
	Dim woking_type
	Dim Working_mnth
	Dim out_date

	woking_type = "CC"
	Working_mnth = Trim(Request("txtWorkingDt"))

    Call SubCreateCommandObject(lgObjComm)	 

    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "usp_c_close_movement"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@work_type"  ,adVarChar,adParamInput,LEN(woking_type), Trim(woking_type))
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarChar,adParamInput,13, gUsrID)
		.Parameters.Append .CreateParameter("@yyyymm"     ,adVarChar,adParamInput,LEN(Working_mnth), Trim(Working_mnth))	    
        .Parameters.Append .CreateParameter("@out_date"   ,adVarChar,adParamOutput,8)				
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
		out_date = lgObjComm.Parameters("@out_date").Value        

        If  IntRetCD <> 1 Then
			out_date = ""
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '¢Ð: Protect system from crashing   
			Response.end
		Else	
			out_date = lgObjComm.Parameters("@out_date").Value
        End If
    Else
        lgErrorStatus = "YES"                                                         '¢Ð: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If

	Call SubCloseCommandObject(lgObjComm)	
	
	Response.Write " <Script Language=vbscript>	 " & vbCr
	Response.Write " With parent				 " & vbCr
	Response.Write "	Call .GenCloseDateInfo() " & vbCr		
	Response.Write " End With					 " & vbCr
	Response.Write " </Script>					 " & vbCr 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizCancel()
	Dim IntRetCD
	Dim strMsg_cd
	Dim woking_type
	Dim Working_mnth
	Dim out_date

	woking_type = "CD"
	Working_mnth = Trim(Request("txtWorkingDt"))

    Call SubCreateCommandObject(lgObjComm)	 

    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "usp_c_close_movement"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@work_type"  ,adVarChar,adParamInput,LEN(woking_type), Trim(woking_type))
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarChar,adParamInput,13, gUsrID)
		.Parameters.Append .CreateParameter("@yyyymm"     ,adVarChar,adParamInput,LEN(Working_mnth), Trim(Working_mnth))
        .Parameters.Append .CreateParameter("@out_date"   ,adVarChar,adParamOutput,8)		
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        If  IntRetCD <> 1 Then
			out_date = ""
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '¢Ð: Protect system from crashing   
			Response.end
		Else
			out_date = lgObjComm.Parameters("@out_date").Value			
        End If
    Else
        lgErrorStatus = "YES"                                                         '¢Ð: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If

	Call SubCloseCommandObject(lgObjComm)	

	Response.Write " <Script Language=vbscript>	 " & vbCr
	Response.Write " With parent                 " & vbCr
	Response.Write "	Call .GenCloseDateInfo() " & vbCr		
	Response.Write " End With					 " & vbCr
	Response.Write " </Script>					 " & vbCr

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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub


%>

<Script Language="VBScript">
    Parent.ExeReflectOk1    
    Select Case "<%=lgOpModeCRUD %>"
      Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
			 Parent.ExeReflectOk
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.ExeReflectOk
          End If   
     End Select    
</Script>
