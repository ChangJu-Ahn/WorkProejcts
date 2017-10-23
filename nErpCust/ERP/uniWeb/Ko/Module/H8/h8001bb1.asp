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

    Dim strfr_pay_yymm
    Dim strto_pay_yymm
    Dim strProv_type
    Dim strPay_cd

    Dim strbas_mm
    Dim strbas_dd

    Dim strprov_mm
    Dim strprov_dd
    Dim strRetro_yn
    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strfr_pay_yymm = Request("txtfr_pay_yymm")
    strto_pay_yymm = Request("txtto_pay_yymm")
    strProv_type = Request("txtProv_cd")
    strbas_mm = Request("txtstd_mm")
    strbas_dd = Request("txtstd_dd")

    strprov_mm = Request("txtprov_mm")
    strprov_dd = Request("txtprov_dd")

    strPay_cd = Request("txtPay_cd")
    if  strPay_cd = "" then
        strPay_cd = "%"
    end if
    strRetro_yn = Request("txtRetro_yn")
    strEmp_no = Request("txtEmp_no")
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if

    With lgObjComm
        .CommandText = "usp_retro_main_pay_calc"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",     adXChar,adParamInput,Len(gUsrId),gUsrId)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_pay_yymm",adXChar,adParamInput,Len(strfr_pay_yymm),strfr_pay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_pay_yymm",adXChar,adParamInput,Len(strto_pay_yymm),strto_pay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type",  adXChar,adParamInput,Len(strProv_type),strProv_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_mm",     adXChar,adParamInput,Len(strbas_mm),strbas_mm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dd",     adXChar,adParamInput,Len(strbas_dd),strbas_dd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_mm",    adXChar,adParamInput,Len(strprov_mm),strprov_mm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dd",    adXChar,adParamInput,Len(strprov_dd),strprov_dd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput,Len(strPay_cd),strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no),strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retro_yn",   adXChar,adParamInput,Len(strRetro_yn),strRetro_yn)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamOutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

'CREATE procedure usp_retro_main_pay_calc (@usr_id           VARCHAR(13),  -- 로그인ID
'                                   		  @fr_pay_yymm      VARCHAR(6),   -- 계산시작월 
'                                   		  @to_pay_yymm      VARCHAR(6),   -- 계산종료월 
'                                          @prov_type        VARCHAR(1),   -- 지급구분(P)
'										  @bas_mm		    VARCHAR(1),   -- 기준월 
'                                          @bas_dd			VARCHAR(2),   -- 기준일 
'										  @prov_mm			VARCHAR(1),   -- 지급월 
'                                          @prov_dd			VARCHAR(2),   -- 지급일 
'                    	                  @para_pay_cd      VARCHAR(1),   -- 급여구분 
'                          	              @para_emp_no      VARCHAR(13),  -- 사번 
'	  									  @msg_cd           VARCHAR(6)	OUTPUT, 
'                   	                      @msg_text         VARCHAR(6)	OUTPUT 
'       			   			             ) AS



        lgObjComm.Execute  ,, adExecuteNoRecords

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