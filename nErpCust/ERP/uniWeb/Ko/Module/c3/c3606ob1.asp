<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
	Dim CallType
	Dim strFromYYYYMM
	Dim strPlantCd
	Dim strItemAcctCd
	Dim strCostCd
	Dim strFromCostCd
	Dim strDstbCostCd
	Dim strAcctcd
	Dim strItemCd
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
	
	Call SubBatch()
		
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
	
	strFromYYYYMM	= "" & Request("txtFromYYYYMM")
	strPlantCd		= "" & Request("txtPlantCd")
	strItemAcctCd	= "" & Request("txtItemAcctCd")
	strCostCd	= "" & Request("txtCostCd1")
	strFromCostCd	= "" & Request("txtCostCd2")
	strDstbCostCd	= "" & Request("txtCostCd3")
	strAcctCd	= "" & Request("txtAcctCd")
	strItemCd	= "" & Request("txtItemCd")
	
	    
End Sub     
'============================================================================================================
' Name : SubBomExplode
' Desc : Query Data from Db
'============================================================================================================
Sub SubBatch()

    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    With lgObjComm
		.CommandText = "usp_c_mfc_allc_info"
		
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymm",	adVarXChar,adParamInput,Len(strFromYYYYMM), strFromYYYYMM)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	adVarXChar,adParamInput,Len(strPlantCd), strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@item_acct",	adVarXChar,adParamInput,Len(strItemAcctCd), strItemAcctCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cost_cd",	adVarXChar,adParamInput,Len(strCostCd), strCostCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_cost_cd",	adVarXChar,adParamInput,Len(strFromCostCd), strFromCostCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@dstb_cost_cd",	adVarXChar,adParamInput,Len(strDstbCostCd), strDstbCostCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@acct_cd",	adVarXChar,adParamInput,Len(strAcctCd), strAcctCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@item_cd",	adVarXChar,adParamInput,Len(strItemCd), strItemCd)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id",	adVarXChar,adParamOutput,13)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With
    
	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		If  IntRetCD <> 1 then
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
	If CInt(btnType) < 4 Then
%>
		parent.frm1.txtSpId.value	= "<%=strSpId%>"
		Call parent.PrevExecOk()
<%
	ElseIF CInt(btnType) > 3  Then
%>		
		parent.frm1.txtSpId.value	= "<%=strSpId%>"
		Call parent.PrintExecOk()
<%
	End If
%>
</Script>