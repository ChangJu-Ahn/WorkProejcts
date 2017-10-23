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
    Call LoadinfTb19029B("B", "H","NOCOOKIE","BB")  
    
    Dim intRetCD
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus = "NO"
    lgErrorPos    = ""                                                           '☜: Set to space

    lgOpModeCRUD = Request("txtMode")

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatch()
    Call SubCloseCommandObject(lgObjComm)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim Maxprice, LngRecs

    Dim strycal_yy_dt
    Dim strycal_yymm_dt
    Dim stryear_yymm_dt
    Dim strProv_type
    Dim strPay_cd
    Dim strEmp_no
    Dim strtax_calc
    Dim strEmp_no2
    Dim DblAllow

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strycal_yy_dt  = Request("txtycal_yy_dt")
    strycal_yymm_dt  = Request("txtycal_yymm_dt")
    stryear_yymm_dt = Request("txtyear_yymm_dt")
    strProv_type    = Request("txtProv_type")
    strtax_calc = Request("txttax_calc")

    strPay_cd = Request("txtPay_cd")
    If  strpay_cd = "" Then
        strpay_cd = "%"
    else
        strpay_cd = strpay_cd & "%"
    End If
    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    else
        strEmp_no = strEmp_no & "%"
    End If

    With lgObjComm
        .CommandText = "usp_hfb050b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",     adXChar,adParamInput,Len(gUsrId),         gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yymm",  adXChar,adParamInput,Len(strycal_yy_dt),  strycal_yy_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm",   adXChar,adParamInput,Len(strycal_yymm_dt),strycal_yymm_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dt_s"  ,adXChar,adParamInput,Len(stryear_yymm_dt),stryear_yymm_dt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type",  adXChar,adParamInput,Len(strProv_type),   strProv_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_pay_cd",adXChar,adParamInput,Len(strPay_cd),      strPay_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adXChar,adParamInput,Len(strEmp_no),      strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@tax_calc_flag",adXChar,adParamInput,Len(strtax_calc),  strtax_calc)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adXChar,adParamOutput,60)

'CREATE procedure usp_hfb050b1(@usr_id           VARCHAR(13), -- 로그인 ID
'                              @year_yymm        VARCHAR(6),  -- 연월차계산월 
'                              @pay_yymm		    VARCHAR(6),  -- 연월차반영월 
'--                              @prov_type        VARCHAR(1),  -- 지급구분 
'                              @para_pay_cd      VARCHAR(1),  -- 급여구분 
'                              @para_emp_no      VARCHAR(13),  -- 사번 
'                              @msg_cd           VARCHAR(6)	OUTPUT, 
'                              @msg_text         VARCHAR(60)	OUTPUT


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

Sub SubBizBatch1()

    Dim Maxprice, LngRecs

    Dim strycal_yy_dt
    Dim strycal_yymm_dt
    Dim strProv_type
    Dim strPay_cd
    Dim strEmp_no
    
    Dim strEmp_no2
    Dim DblAllow

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strycal_yy_dt  = Replace(Request("txtycal_yy_dt"),"-", "")
    strycal_yymm_dt  = Replace(Request("txtycal_yymm_dt"),"-", "")
    strProv_type    = Request("txtProv_type")

    strPay_cd = Request("txtPay_cd")
    If  strpay_cd = "" Then
        strpay_cd = "%"
    else
        strpay_cd = strpay_cd & "%"
    End If
    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    else
        strEmp_no = strEmp_no & "%"
    End If

    lgStrSQL = " DELETE hdf030t FROM hdf030t a, hdf020t b "
    lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no"
    lgStrSQL = lgStrSQL & "   AND a.emp_no LIKE  " & FilterVar(stremp_no, "''", "S") & ""
    lgStrSQL = lgStrSQL & "   AND b.pay_cd LIKE  " & FilterVar(strpay_cd, "''", "S") & ""
    lgStrSQL = lgStrSQL & "   AND a.allow_cd = " & FilterVar("P13", "''", "S") & ""

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

    lgStrSQL = " SELECT a.emp_no, ISNULL(SUM(ISNULL(a.tot_amt,0)),0) "
    lgStrSQL = lgStrSQL & " FROM hfb020t a, hdf020t b "
    lgStrSQL = lgStrSQL & "WHERE a.emp_no = b.emp_no "
    lgStrSQL = lgStrSQL & "  AND a.emp_no LIKE  " & FilterVar(stremp_no, "''", "S") & ""
    lgStrSQL = lgStrSQL & "  AND b.pay_cd LIKE  " & FilterVar(strpay_cd, "''", "S") & ""
    lgStrSQL = lgStrSQL & "  AND a.year_yymm =  " & FilterVar(strycal_yy_dt , "''", "S") & ""
    lgStrSQL = lgStrSQL & "GROUP BY a.emp_no "

    IF  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true then

        strycal_yymm_dt = Request("txtycal_yymm_dt") & "-01"

        Do While Not lgObjRs.EOF
            strEmp_no2 = ConvSPChars(lgObjRs(0))
            DblAllow = ConvSPChars(lgObjRs(1))

            lgStrSQL = " INSERT INTO hdf030t VALUES( "
            lgStrSQL = lgStrSQL & FilterVar(strEmp_no2, "''", "S") & "," 
            lgStrSQL = lgStrSQL & "" & FilterVar("P13", "''", "S") & "" & "," 
            lgStrSQL = lgStrSQL & FilterVar(DblAllow,"0","D") & "," 
            lgStrSQL = lgStrSQL & "" & FilterVar("Y", "''", "S") & " " & "," 
            lgStrSQL = lgStrSQL & FilterVar(strycal_yymm_dt, "''", "S") & "," 
            lgStrSQL = lgStrSQL & FilterVar(strycal_yymm_dt, "''", "S") & "," 
            lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
            lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
            lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
            lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
            lgStrSQL = lgStrSQL & " ) "

            lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
            If  CheckSYSTEMError(Err,True) = True Then
                Call DisplayMsgBox("800414", vbInformation, "", "", I_MKSCRIPT)
                exit Sub
            end if

            lgStrSQL = " UPDATE hfb020t SET "
            lgStrSQL = lgStrSQL & " year_flag = " & FilterVar("Y", "''", "S") & "  "
            lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(strEmp_no2, "''", "S")
            lgStrSQL = lgStrSQL & "   AND year_type IN (" & FilterVar("1", "''", "S") & " , " & FilterVar("2", "''", "S") & ") "
            lgStrSQL = lgStrSQL & "   AND year_yymm = " & FilterVar(strycal_yy_dt, "''", "S")

            lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
            If  CheckSYSTEMError(Err,True) = True Then
                Call DisplayMsgBox("800414", vbInformation, "", "", I_MKSCRIPT)
                exit Sub
            end if

            lgObjRs.MoveNext
        Loop
        IntRetCD = 1
    else
        IntRetCD = -1
    end if

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
