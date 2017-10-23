<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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
    Dim strdilig_cd
    Dim strtax_calc
    Dim strPay_cd
    Dim strallow_cd
    Dim strRetire_flag, strRetire_stdt, strRetire_enddt

    Dim strEmp_no

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    stryear_yymm = Request("txtyear_yymm")
    stryear_type = Request("txtyear_type")
    strdilig_cd = Request("txtdilig_cd")
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

    strRetire_flag = "Y"
    strRetire_stdt = Request("txtRetire_stdt")
    strRetire_enddt = Request("txtRetire_enddt")

    With lgObjComm
        
        IF  stryear_type  = "1" THEN  '연차이면 

        	IF  strallow_cd <> "" THEN
        		IF  strtax_calc = "Y" THEN
        		    .CommandText = "usp_hfb010b1"
        		    .CommandType = adCmdStoredProc
                    'll_return = SQLCA.usp_hfb010b1(ida.user_id, ls_year_yymm + ls_year_yymm_dt, &
			        '              ls_year_type + ls_pay_cd, ls_allow_cd, ls_dilig_cd, ls_emp_no)
                ELSE
                    .CommandText = "usp_hfb011b1"
        		    .CommandType = adCmdStoredProc
                    'll_return = SQLCA.usp_hfb011b1(ida.user_id, ls_year_yymm + ls_year_yymm_dt, &
			        '              ls_year_type, ls_allow_cd, ls_dilig_cd, ls_pay_cd, ls_emp_no)
                END IF

        	END IF

        ElseIF  stryear_type  = "2" THEN  '월차이면 

        	IF  strallow_cd <> "" THEN
        		IF  strtax_calc = "Y" THEN
        		    .CommandText = "usp_hfb030b1"
        		    .CommandType = adCmdStoredProc
        	   	    'll_return = SQLCA.usp_hfb030b1(ida.user_id, ls_year_yymm + ls_year_yymm_dt, & 
        			'                               ls_year_type + ls_pay_cd, ls_allow_cd, ls_dilig_cd, ls_emp_no)
                ELSE
                    .CommandText = "usp_hfb031b1"
        		    .CommandType = adCmdStoredProc
        	   	    'll_return = SQLCA.usp_hfb031b1(ida.user_id, ls_year_yymm + ls_year_yymm_dt, & 
        		    '	                            ls_year_type, ls_allow_cd, ls_dilig_cd, ls_pay_cd, ls_emp_no)
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
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@dilig_cd"   ,adXChar,adParamInput,Len(strdilig_cd),    strdilig_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no),      strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_flag"  ,adXChar,adParamInput,Len(strRetire_flag), strRetire_flag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_stdt_s"  ,adXChar,adParamInput,Len(strRetire_stdt), strRetire_stdt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_enddt_s" ,adXChar,adParamInput,Len(strRetire_enddt),strRetire_enddt)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

'CREATE procedure usp_hfb010b1(@usr_id           VARCHAR(13),    /* 로그인 ID */
'                              @year_yymm        VARCHAR(6),     -- 연월차년월 
'                              @prov_dt_s        VARCHAR(8),     --연월차지급일 
'                              @year_type        VARCHAR(1),     -- 연월차구분 
'							  @para_pay_cd      VARCHAR(1),     -- 급여구분 
'                              @allow_cd         VARCHAR(3),     /* 연차수당코드 */
'                              @dilig_cd         VARCHAR(2),     /* 연차근태코드 */
'                              @para_emp_no      VARCHAR(13),     /* 사번 */
'						      @msg_cd      	    VARCHAR(6)	OUTPUT,     -- Error Message Code 
'							  @msg_text         VARCHAR(6)	OUTPUT      -- Error Message

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