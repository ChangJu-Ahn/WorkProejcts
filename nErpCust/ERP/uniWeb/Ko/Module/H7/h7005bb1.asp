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

    lgOpModeCRUD      = Request("txtMode")
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         
             Call SubCreateCommandObject(lgObjComm)
			 Call SubBizBatch()
             Call SubCloseCommandObject(lgObjComm)
    
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubOpenDB(lgObjConn)
             Call SubBizSaveMultiDelete()
             Call SubCloseDB(lgObjConn)
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strBonus_yymm_dt
    Dim strProv_type
    Dim strPay_cd
    Dim strTax_strt_dt
    Dim strTax_end_dt

    Dim strPay_type
    Dim strCalcu_type
    Dim strRetire_flag
    Dim strSave_flag
    Dim strLoan_flag

    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strBonus_yymm_dt = Request("txtBonus_yymm_dt")
    strProv_type = Request("txtProv_type")
    strPay_cd = Request("txtPay_cd")
    if  strPay_cd = "" then
        strPay_cd = "%"
    end if

    strTax_strt_dt = Request("txtTax_strt_dt")
    strTax_end_dt = Request("txtTax_end_dt")

    strEmp_no = Request("txtEmp_no")
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if

    strPay_type = Request("txtPay_type")
    strCalcu_type = Request("txtCalcu_type")
    strRetire_flag = Request("txtRetire_flag")
    strSave_flag = Request("txtSave_flag")
    strLoan_flag = Request("txtLoan_flag")

    With lgObjComm
        .CommandText = "usp_hea050b1"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput,Len(gUsrId), gUsrId)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bonus_cd"   ,adXChar,adParamInput,Len("000"), "000")
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adXChar,adParamInput,Len(strBonus_yymm_dt), strBonus_yymm_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no), strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adXChar,adParamInput,Len(strProv_type), strProv_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput,Len(strPay_cd), strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@calcu_type" ,adXChar,adParamInput,Len(strCalcu_type), strCalcu_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@tax_strt_dt",adXChar,adParamInput,Len(strTax_strt_dt), strTax_strt_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@tax_end_dt" ,adXChar,adParamInput,Len(strTax_end_dt), strTax_end_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_type"   ,adXChar,adParamInput,Len(strPay_type), strPay_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_type",adXChar,adParamInput,Len(strRetire_flag), strRetire_flag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@save_flag"  ,adXChar,adParamInput,Len(strSave_flag), strSave_flag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@loan_flag"  ,adXChar,adParamInput,Len(strLoan_flag), strLoan_flag)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)
        
        lgObjComm.Execute ,, adExecuteNoRecords

'CREATE procedure usp_hea050b1 (@usr_id           VARCHAR(13), /* 로그인 ID */
'                               @bonus_cd         VARCHAR(3),  /* 000 */
'                               @pay_yymm         VARCHAR(6),  /* 상여년월 */
'                               @para_emp_no      VARCHAR(13), /* 사번 */
'                               @prov_type        VARCHAR(1),  /* 지급구분 */
'                               @para_pay_cd      VARCHAR(1),  /* 급여구분 */
'                               @calcu_type       VARCHAR(1),  /* 세액계산여부 */
'                               @tax_strt_dt      VARCHAR(6),  /* 세금계산 시작일 */
'                               @tax_end_dt       VARCHAR(6),  /* 세금계산 종료일 */
'                               @pay_type         VARCHAR(1),  /* 중도정산시 급/상여포함여부(1, 2) */
'                               @retire_type      VARCHAR(1),  /* 퇴직금 포함여부(Y, N) */
'                               @save_flag        VARCHAR(1),  /* 상여저축계산여부(Y, N) */
'                               @loan_flag        VARCHAR(1),  /* 상여대부상환계산여부(Y, N) */'


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
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error statu


    Dim strBonus_yymm_dt
    Dim strProv_type
    Dim strPay_cd
    Dim strEmp_no


    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strBonus_yymm_dt = Request("txtBonus_yymm_dt")
    strProv_type = Request("txtProv_type")
    strPay_cd = Request("txtPay_cd")
    strEmp_no = Request("txtEmp_no")

'Response.Write " -strBonus_yymm_dt:" & strBonus_yymm_dt
'Response.Write " -strProv_type:" & strProv_type
'Response.Write " -strPay_cd:" & strPay_cd
'Response.Write " -strEmp_no:" & strEmp_no
'Response.End

    lgStrSQL = "DELETE  hdf070t"
    lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar( strBonus_yymm_dt ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar( strProv_type ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND pay_cd like " & FilterVar( strPay_cd ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no like " & FilterVar( strEmp_no ,"''", "S")


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

    lgStrSQL = "DELETE  hdf041t"
    lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar( strBonus_yymm_dt ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar( strProv_type ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no like " & FilterVar( strEmp_no ,"''", "S")    
    lgStrSQL = lgStrSQL & "   AND emp_no in (select emp_no from hdf020t where pay_cd like " & FilterVar( strPay_cd ,"''", "S") & ")"
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)


    lgStrSQL = "DELETE  hdf060t"
    lgStrSQL = lgStrSQL & " WHERE sub_yymm = " & FilterVar( strBonus_yymm_dt ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_type = " & FilterVar( strProv_type ,"''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no like " & FilterVar( strEmp_no ,"''", "S")     
    lgStrSQL = lgStrSQL & "   AND emp_no in (select emp_no from hdf020t where pay_cd like " & FilterVar( strPay_cd ,"''", "S") & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

'select * from  HDC020T  --월저축불입내역 
'select * from  HDD020T  --월대부상환내역 
'select * from HFB030T  --연월차기준금액내역 
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub


%>

<Script Language="VBScript">

    With Parent
		Select Case "<%=lgOpModeCRUD %>"
		    Case "<%=UID_M0001%>"
			    IF Trim("<%=lgErrorStatus%>") = "NO" AND "<%=CInt(intRetCD)%>" >= 0 Then
			        .ExeReflectOk
			    Else
			        .ExeReflectNo
			    End If
		    Case "<%=UID_M0003%>"
		        If Trim("<%=lgErrorStatus%>") = "NO" Then
			        .ExeDeleteOk
			    Else
			        .ExeReflectNo
			    End If  
		End Select
    End with	   
</Script>