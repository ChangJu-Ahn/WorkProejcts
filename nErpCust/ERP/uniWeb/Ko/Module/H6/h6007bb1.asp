<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Dim intRetCD

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatch()
    Call SubCloseCommandObject(lgObjComm)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim iDx

    Dim strDt_from
    Dim strDt_to
    Dim strPay_yymm
    Dim strPay_cd
    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strDt_from = Request("txtDt_from")
    strDt_to = Request("txtDt_to")
    strPay_yymm = Request("txtPay_yymm")
    strPay_cd = Request("txtPay_cd")
    strEmp_no = Request("txtEmp_no")
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if
    if  strPay_cd = "" then
        strPay_cd = "%"
    end if   
   
    With lgObjComm
        .CommandText = "usp_hba100b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput,Len(gUsrID), gUsrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_dt"    ,adXChar,adParamInput,Len(strDt_from), strDt_from)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_dt"      ,adXChar,adParamInput,Len(strDt_to), strDt_to)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adXChar,adParamInput,Len(strPay_yymm), strPay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput,Len(strPay_cd), strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no), strEmp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

'CREATE procedure usp_hba100b1(@usr_id           VARCHAR(13), -- 로그인 ID
'                              @from_dt          VARCHAR(8),  -- 시작월 
'                              @to_dt            VARCHAR(8),  -- 종료월 
'                              @pay_yymm		    VARCHAR(6),  -- 학자금반영월 
'                              --@prov_type        VARCHAR(1),  -- 지급구분 
'                              @para_pay_cd      VARCHAR(1),  -- 급여구분 
'                              @para_emp_no      VARCHAR(13),  -- 사번 
'                              @msg_cd           VARCHAR(6), 
'                              @msg_text         VARCHAR(60) 

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