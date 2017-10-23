<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    Dim intRetCD

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    call LoadBasisGlobalInf()

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space

     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strMsg_cd
    Dim strMsg_text

    Dim strPay_yymm
    Dim strPay_cd
    Dim strEmp_no

    Dim strOrg_cd

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strOrg_cd    = Request("txtbiz_area")
    strPay_yymm  = Request("txtPay_yymm")
    strPay_cd    = Request("txtPay_cd")
    strEmp_no    = Request("txtEmp_no")
    
    If  strEmp_no = "" Then
        strEmp_no = "%"
    End If

    If  strPay_cd = "" Then
        strPay_cd = "%"
    End If

    With lgObjComm
        .CommandText = "usp_hca120b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_cd"     ,adVarXChar,adParamInput,Len(strOrg_cd), strOrg_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput,Len(gUsrID), gUsrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adVarXChar,adParamInput,Len(strPay_yymm), strPay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adVarXChar,adParamInput,Len(strEmp_no), strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adVarXChar,adParamInput,Len(strPay_cd), strPay_cd)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamOutput,6)
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status
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
