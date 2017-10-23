<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
	Dim CallType
	Dim strPlantCd
	Dim strFrItemCd
	Dim strToItemCd
	Dim btnType
	Dim strSpId

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

	Call SubOpenDB(lgObjConn)       
	Call SubCreateCommandObject(lgObjComm)
	Call SubMakeParameter()
	Call SubBomExplode()
		
	Call SubCloseCommandObject(lgObjComm)
	Call SubCloseDB(lgObjConn)      
 
'============================================================================================================
' Name : SubMakeParameter
' Desc : Make SP Parameter
'============================================================================================================
Sub SubMakeParameter()
	
	If CInt(Request("txtMode")) <> UID_M0001 Then
		Response.End 
	End If
	
	strPlantCd		= "" & Request("txtPlantCd")
	strFrItemCd		= "" & Request("txtFrItemCd")
	strToItemCd		= "" & Request("txtToItemCd")
	
	btnType			= Request("BtnType")
	    
End Sub     
'============================================================================================================
' Name : SubBomExplode
' Desc : Query Data from Db
'============================================================================================================
Sub SubBomExplode()

    Dim strMsg_cd
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    With lgObjComm
        .CommandText = "usp_c_std_bom_main"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	adVarChar,adParamInput,Len(strPlantCd), strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_item_cd",	adVarChar,adParamInput,Len(strFrItemCd), strFrItemCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_item_cd",	adVarChar,adParamInput,Len(strToItemCd), strToItemCd)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	adVarChar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id",	adVarChar,adParamOutput,13)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With
    
	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		If  IntRetCD <> 1 then
			strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
		            
			If strMsg_cd <> "" Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
				Response.End
			End If
			Exit Sub
		Else
			strSpId = lgObjComm.Parameters("@sp_id").Value
		End if
	Else           
		Call SvrMsgBox(Err.Description, VBInformation, I_MKSCRIPT)
		Response.End
	End if
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
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language=vbscript>
	
<% 		
	If btnType = 0 Then
%>
		parent.frm1.txtSpId.value	= "<%=strSpId%>"
		Call parent.PrevExecOk()
<%
	ElseIF btnType = 1  Then
%>		
		parent.frm1.txtSpId.value	= "<%=strSpId%>"
		Call parent.PrintExecOk()
<%
	End If
%>
</Script>