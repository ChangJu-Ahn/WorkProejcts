<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("B", "H","NOCOOKIE","BB")                                                                     '☜: Clear Error status
    Dim intRetCD

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

    Dim stryear_yymm
    Dim stryear_type
    Dim strtax_calc
    Dim strPay_cd
    Dim strallow_cd
    Dim strRetire_flag, strRetire_stdt, strRetire_enddt
    Dim strdilig_cd
    
    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    stryear_yymm = Request("txtyear_yymm")
    stryear_type = Request("txtyear_type")
    strtax_calc = Request("txttax_calc")

    strPay_cd = Request("txtPay_cd")
    if  strPay_cd = "" then
        strPay_cd = "%"
    end if
    strEmp_no = Request("txtEmp_no")
    if  strEmp_no = "" then
        strEmp_no = "%"
    end if
    strallow_cd = Request("txtallow_cd")
    strdilig_cd = Request("txtdilig_cd")
    strRetire_flag = Request("txtRetireFlag")
    strRetire_stdt = Request("txtRetire_stdt")
    strRetire_enddt = Request("txtRetire_enddt")

    With lgObjComm
        
        IF  stryear_type  = "1" THEN  '연차이면 

        	IF  strallow_cd <> "" THEN
        		.CommandText = "usp_hfb010b1"
        		.CommandType = adCmdStoredProc
        	END IF

        ElseIF  stryear_type  = "2" THEN  '월차이면 

        	IF  strallow_cd <> "" THEN
        		IF  strtax_calc = "Y" THEN
        		    .CommandText = "usp_hfb030b1"
        		    .CommandType = adCmdStoredProc
                ELSE
                    .CommandText = "usp_hfb031b1"
        		    .CommandType = adCmdStoredProc
                END IF

        	END IF
        ELSE
            IntRetCD = -1
            Exit Sub
        END IF

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adXChar,adParamInput,Len(gUsrId),         gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yymm"  ,adXChar,adParamInput,Len(stryear_yymm),   stryear_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_type"  ,adXChar,adParamInput,Len(stryear_type),   stryear_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput,Len(strPay_cd),      strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@allow_cd"   ,adXChar,adParamInput,Len(strallow_cd),    strallow_cd)
        
        IF  stryear_type  = "2" THEN  '월차이면	    
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@dilig_cd"   ,adXChar,adParamInput,Len(strdilig_cd),    strdilig_cd)
		End If 
		
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no),      strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_flag"  ,adXChar,adParamInput,Len(strRetire_flag), strRetire_flag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_stdt_s"  ,adXChar,adParamInput,Len(strRetire_stdt), strRetire_stdt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_enddt_s" ,adXChar,adParamInput,Len(strRetire_enddt),strRetire_enddt)
	    
        IF  stryear_type  = "1" THEN  '연차이면 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@tax_yn"  ,adXChar,adParamInput,Len(strtax_calc), strtax_calc)
		End If 
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