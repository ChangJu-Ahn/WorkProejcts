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

    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPC3G055
	Dim CostClsYYYYMM
	Dim intRetCd
	Dim CostRefYYYYMM
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	intRetCd = CommonQueryRs("isnull(max(YYYYMM)," & FilterVar("X", "''", "S") & " )","c_close_status","" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) = "X"  then
		CostRefYYYYMM = ""
	else
		CostRefYYYYMM = Trim(Replace(lgF0,Chr(11),""))	  
	end if   


    IF CostRefYYYYMM <> "" Then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostRefDt.Year = """ & ConvSPChars(Mid(CostRefYYYYMM,1,4))   & """" & vbCr
		Response.Write "	.frm1.txtCostRefDt.Month = """ & ConvSPChars(Mid(CostRefYYYYMM,5,2))   & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	ELSE
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostRefDt.text = """ & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	END IF	

	set lgF0 = nothing
		
    Set iPC3G055 = Server.CreateObject("PC3G055.cCMngCostClosingSvr")

    If CheckSYSTEMError(Err, True) = True Then
	   SetErrorStatus()						
       Exit Sub
    End If    

	
    Call iPC3G055.C_MANAGE_COST_CLOSING_SVR (gStrGloBalCollection,"C",CostClsYYYYMM)		
		
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPC3G055 = Nothing
       SetErrorStatus()
       Exit Sub
    End If    
    
    Set iPC3G055 = Nothing

   

    IF CostClsYYYYMM <> "" Then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostClsDt.Year = """ & ConvSPChars(Mid(CostClsYYYYMM,1,4))   & """" & vbCr
		Response.Write "	.frm1.txtCostClsDt.Month = """ & ConvSPChars(Mid(CostClsYYYYMM,5,2))   & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	ELSE
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostClsDt.text = """ & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	END IF	
	
	
    '---------- Developer Coding part (End)   ---------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizCancel()

    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear                                                                            '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPC3G055
	Dim CostClsYYYYMM
	Dim intRetCd
	Dim CostRefYYYYMM
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	intRetCd = CommonQueryRs("isnull(max(YYYYMM)," & FilterVar("X", "''", "S") & " )","c_close_status","" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) = "X"  then
		CostRefYYYYMM = ""
	else
		CostRefYYYYMM = Trim(Replace(lgF0,Chr(11),""))	  
	end if   


    IF CostRefYYYYMM <> "" Then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostRefDt.Year = """ & ConvSPChars(Mid(CostRefYYYYMM,1,4))   & """" & vbCr
		Response.Write "	.frm1.txtCostRefDt.Month = """ & ConvSPChars(Mid(CostRefYYYYMM,5,2))   & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	ELSE
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostRefDt.text = """ & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	END IF	

	set lgF0 = nothing   

    Set iPC3G055 = Server.CreateObject("PC3G055.cCMngCostClosingSvr")

    If CheckSYSTEMError(Err, True) = True Then
	   SetErrorStatus()						
       Exit Sub
    End If    

	
    Call iPC3G055.C_MANAGE_COST_CLOSING_SVR (gStrGloBalCollection,"D",CostClsYYYYMM)		
		
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPC3G055 = Nothing
       SetErrorStatus()
       Exit Sub
    End If    
    
    Set iPC3G055 = Nothing
    

    IF CostClsYYYYMM <> "" Then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostClsDt.Year = """ & ConvSPChars(Mid(CostClsYYYYMM,1,4))   & """" & vbCr
		Response.Write "	.frm1.txtCostClsDt.Month = """ & ConvSPChars(Mid(CostClsYYYYMM,5,2))   & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	ELSE
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtCostClsDt.text = """ & """" & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr 
	END IF

    
 
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
