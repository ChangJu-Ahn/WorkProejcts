<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Dim intRetCD

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD      = Request("txtMode") 
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
	
    Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0006)                                                         '☜: Query
			 Call SubBizBatch()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizBatchDelete()
    End Select
         
    Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strpay_yymm
    Dim strbas_dt
    Dim strprov_dt
    Dim strProv_cd
    Dim strPay_cd
    Dim strChk_pay_cd
    Dim strEmp_no
    Dim strStand
    Dim strStand_dt
    Dim Li_yes  '12월 월차수당이 12월급여에 포함되는 OPTION입니까?
    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strpay_yymm = Request("txtpay_yymm_dt")
    strbas_dt   = Request("txtbas_dt")
    strprov_dt  = Request("txtprov_dt")
    strProv_cd  = Request("txtProv_cd")
    strStand    = Request("txtStand")
    Li_yes      = Request("txtLi_yes")

    strPay_cd = Request("txtPay_cd")
    strChk_pay_cd = Request("txtPay_cd")
    If  strpay_cd = "" Then
        strpay_cd = "%"
    else
        strpay_cd = strpay_cd
    End If
    If  strchk_pay_cd = "" Then
        strchk_pay_cd = "*"
    end if
    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    End If


'call svrmsgbox(lgOpModeCRUD &"/"& gUsrId &"/"&  stryear_yymm &"/"& stryear_type & "/"& strallow_cd &"/"& strEmp_no , vbinformation,i_mkscript) 


    With lgObjComm
        .CommandText = "usp_main_pay_calc"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput, Len(gUsrId),      gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adXChar,adParamInput, Len(strPay_yymm), strPay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adXChar,adParamInput, Len(strProv_cd),  strProv_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dt_s"   ,adXChar,adParamInput, Len(strBas_dt),   strBas_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dt_s"  ,adXChar,adParamInput, Len(strProv_dt),  strProv_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput, Len(strPay_cd),   strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput, Len(strEmp_no),   strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@li_yes"     ,adXChar,adParamInput, Len(Li_yes),      Li_yes)

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)

        lgObjComm.Execute ,, adExecuteNoRecords

    End With


    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value        
        if  IntRetCD < 0 then
            strMsg_cd = Trim(lgObjComm.Parameters("@msg_cd").Value)
            strMsg_text = Trim(lgObjComm.Parameters("@msg_text").Value)

            ObjectContext.SetAbort
                        
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
' Name : SubBizBatchDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizBatchDelete()

    Dim strpay_yymm
    Dim strprov_dt
    Dim strProv_cd
    Dim strPay_cd
    Dim strChk_pay_cd
    Dim strMsg_cd
    Dim strMsg_text
    Dim strEmp_no
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strpay_yymm = Request("txtpay_yymm_dt")
    strprov_dt  = Request("txtprov_dt")
    strProv_cd  = Request("txtProv_cd")


    strPay_cd = Request("txtPay_cd")
    strChk_pay_cd = Request("txtPay_cd")

    If  strpay_cd = "" Then
        strpay_cd = "%"
    else
        strpay_cd = strpay_cd
    End If

    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    End If
    
'call svrmsgbox(strpay_cd & "/" &strchk_pay_cd , vbinformation,i_mkscript)

       			   			       
    With lgObjComm
        .CommandText = "usp_main_pay_delete"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adXChar,adParamInput, Len(strPay_yymm), strPay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adXChar,adParamInput, Len(strProv_cd),  strProv_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_cd"     ,adXChar,adParamInput, Len(strPay_cd),   strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no"     ,adXChar,adParamInput, Len(strEmp_no),   strEmp_no)
	    
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)

        lgObjComm.Execute ,, adExecuteNoRecords

    End With

        
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value        
        if  IntRetCD < 0 then
            strMsg_cd = Trim(lgObjComm.Parameters("@msg_cd").Value)
            strMsg_text = Trim(lgObjComm.Parameters("@msg_text").Value)
            
            ObjectContext.SetAbort
            
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
