<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
    
    Call LoadBasisGlobalInf()

    Dim intRetCD

    Call HideStatusWnd                                                               '��: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space

     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strYear, strBas_dt

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    strYear      = Replace(Request("txtYear_yy"),"''", "")
    strBas_dt    = Replace(Request("txtYear_yy") + "1231","''", "")

    With lgObjComm
        .CommandText = "usp_hfa030b1"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yy "   ,adXChar,adParamInput, 4, strYear)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dt"     ,adXChar,adParamInput, 8, strBas_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput,13, gUsrId)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

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
    End if

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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status
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
               .BundleCreateOk
           Else
               .BundleCreateNo
           End If
      End with
   End If   
       
</Script>	
