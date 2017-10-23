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

    Dim strBonus_yymm_dt
    Dim strProv_type
    Dim strBas_dt
    Dim strProv_dt
    Dim strPay_cd
    Dim strDilig_strt_dt
    Dim strDilig_end_dt
    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strBonus_yymm_dt = Request("txtBonus_yymm_dt")
    strProv_type = Request("txtProv_type")
    strBas_dt = Request("txtBas_dt")
    strProv_dt = Request("txtProv_dt")

    strPay_cd = Request("txtPay_cd")
    if  strPay_cd = "" then
        strPay_cd = "%"
    end if

    strDilig_strt_dt = Request("txtDilig_strt_dt")
    if  strDilig_strt_dt = "" then
        strDilig_strt_dt = "25001231"
    end if
    strDilig_end_dt = Request("txtDilig_end_dt")
    if  strDilig_end_dt = "" then
        strDilig_end_dt = "25001231"
    end if
    strEmp_no = Request("txtEmp_no")
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if

    With lgObjComm
        .CommandText = "usp_hea040b1"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"         ,adXChar,adParamInput,Len(gUsrId), gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bonus_yymm"     ,adXChar,adParamInput,Len(strBonus_yymm_dt), strBonus_yymm_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"      ,adXChar,adParamInput,Len(strProv_type), strProv_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd"    ,adXChar,adParamInput,Len(strPay_cd), strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dt_s"       ,adXChar,adParamInput,Len(strBas_dt), strBas_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dt_s"      ,adXChar,adParamInput,Len(strProv_dt), strProv_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@dilig_strt_s"   ,adXChar,adParamInput,Len(strDilig_strt_dt), strDilig_strt_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@dilig_end_s"    ,adXChar,adParamInput,Len(strDilig_end_dt), strDilig_end_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no"    ,adXChar,adParamInput,Len(strEmp_no), strEmp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

'CREATE procedure usp_hea040b1(@usr_id           VARCHAR(13), /* 로그인 ID */
'                              @bonus_yymm       VARCHAR(6),  /* 상여년월 */
'                              @prov_type        VARCHAR(1),  /* 지급구분 */
'							  @para_pay_cd      VARCHAR(1),  --급여구분 
'                              @bas_dt_s         VARCHAR(8),  /* 기준일 */
'                              @prov_dt_s        VARCHAR(8),  /* 지급일 */
'                              @dilig_strt_s     VARCHAR(8),  /* 근태시작일(30001231) */
'                              @dilig_end_s      VARCHAR(8),  /* 근태종료일(30001231) */
'                              @para_emp_no      VARCHAR(13),  /* 사번 */'
'							  @msg_cd      	 	VARCHAR(6)	OUTPUT,     -- Error Message Code 
'   						      @msg_text         VARCHAR(60)	OUTPUT      -- Error Message 


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