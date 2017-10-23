<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")

    Dim intRetCD

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space

    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()
    Dim strEmp_no
    Dim strRetire_dt
    Dim strCalcu_logic
    Dim strPay_logic
    Dim strMsg_cd,strMsg_text
    Dim IntRetCD
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strEmp_no       = Trim(lgKeyStream(1))
    strRetire_dt    = UNIConvDateToYYYYMMDD(lgKeyStream(2),gServerDateFormat,"")
    strCalcu_logic  = Trim(lgKeyStream(3))
    strPay_logic    = Trim(lgKeyStream(4))
    
    With lgObjComm
        .CommandText = "usp_hga050b1_1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"        ,adVarXChar,adParamInput, 13, gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_cd"     ,adVarXChar,adParamInput, 3,  "R01")
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cal_logic"     ,adVarXChar,adParamInput, 1,  strCalcu_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_logic"     ,adVarXChar,adParamInput, 1,  strPay_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_strt_s" ,adVarXChar,adParamInput, 8,  strRetire_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_end_s"  ,adVarXChar,adParamInput, 8,  strRetire_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no"   ,adVarXChar,adParamInput, 13, strEmp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"        ,adVarXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"      ,adVarXChar,adParamOutput,60)

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
