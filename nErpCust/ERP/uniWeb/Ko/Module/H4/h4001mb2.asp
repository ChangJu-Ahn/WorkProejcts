<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    Dim intRetCD
    call LoadBasisGlobalInf()
    
    Call HideStatusWnd                                                               '⑿: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '⑿: Set to space
	
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim txtStrYear,txtEndYear
    Dim strBA_cd
    Dim strWork
    Dim strMsg_cd
    Dim strMsg_text
    Dim txtDays
    
    strOrg_cd = ""
    txtStrYear  = Replace(Request("txtStrYear"),gComDateType, "")
    txtEndYear  = Replace(Request("txtEndYear"),gComDateType, "")
    txtDays = Request("txtDays")
    
    strBA_cd    = Request("txtBA_cd")
    strWork = Request("txtWork")

    With lgObjComm
        .CommandText = "usp_calendar"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@strYYMM"     ,adXChar,adParamInput, 6, txtStrYear)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@endYYMM"     ,adXChar,adParamInput, 6, txtEndYear)	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_cd"     ,adXChar,adParamInput,Len(strBA_cd), strBA_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@wk_type"    ,adXChar,adParamInput, 1, strWork)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@selDay"    ,adXChar,adParamInput, 1, txtDays)	'林5老 利侩 夸老(N:林5老 利侩救窃,1:岿~6:配) - 2006.04.24
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id"    ,adXChar,adParamInput,Len(gUsrId), gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)
	   
	    lgObjComm.Execute ,, adExecuteNoRecords
        
	End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        if  IntRetCD < 0 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            IntRetCD = -1
            Exit Sub
        else
            IntRetCD = 1
        end if
    Else           
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End If
    
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
    lgErrorStatus     = "YES"                                                         '⑿: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '⑿: Protect system from crashing
    Err.Clear                                                                         '⑿: Clear Error status
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

<Script Language="VBScript">

   If Trim("<%=lgErrorStatus%>") = "NO" Then
      With Parent
           IF  "<%=CInt(intRetCD)%>" >= 0 Then
               .ExeReflectOk
           Else
               .ExeReflectNo
           End If
      End with
   End If   
       
</Script>	
