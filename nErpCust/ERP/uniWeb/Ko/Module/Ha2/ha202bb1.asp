<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")                                                                      '¢Ð: Clear Error status
	
	Dim intRetCD

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

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

    Dim strcalcu_logic
    Dim strpay_logic
    Dim strEmp_no

    Dim strbas_dt
    Dim strretro_bas_dt
    
    Dim intCnt1
    Dim intCnt2
    Dim emp_no, strRetire_yymm

    Dim strMsg_cd
    Dim strMsg_text    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status


    strbas_dt  = Request("txtbas_dt")
    strretro_bas_dt  = Request("txtretro_bas_dt")
    
    strcalcu_logic		= Request("txtcalcu_logic")
    strpay_logic		= Request("txtpay_logic")

    strEmp_no = Trim(Request("txtEmp_no"))
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if

'    if  strpay_logic = "D" then

'        strRetire_yymm = Mid(lgObjRs("retire_yymm"), 1, 6)                
  
 '       If 	CommonQueryRs(" COUNT(*) "," HCA090T "," wk_yymm =  " & FilterVar(strRetire_yymm, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
  '          intCnt1 = CInt(Replace(lgF0, Chr(11), ""))
   '     end if
'
 '       if  intCnt1 = 0 then
  '          Call DisplayMsgBox("800377", vbInformation, "" , "", I_MKSCRIPT)
   '         Call SetErrorStatus
    '        Call SubCloseDB(lgObjConn)
     '       Exit Sub
      '  end if

 '   end if

    With lgObjComm
        .CommandText = "usp_hga060b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,  adXChar,adParamInput,Len(gUsrId), gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_cd",    adXChar,adParamInput,Len("R01"), "R01")
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cal_logic",    adXChar,adParamInput,Len(strcalcu_logic),   strcalcu_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_logic",    adXChar,adParamInput,Len(strPay_logic),     strPay_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dt_s",		adXChar,adParamInput,Len(strbas_dt),	strbas_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retro_bas_dt_s", adXChar,adParamInput,Len(strRetro_bas_dt), strRetro_bas_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",  adXChar,adParamInput,Len(strEmp_no),        strEmp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,  adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,  adXChar,adParamOutput,60)

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
