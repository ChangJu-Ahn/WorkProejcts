<%@ LANGUAGE=VBSCript %>
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
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")                                                                      '☜: Clear Error status
	
	Dim intRetCD

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD      = Request("txtMode")
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

    Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
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

    Dim strretire_yyyy
    Dim strretire_strt_dt,strretire_strt_dt1
    Dim strretire_end_dt,strretire_end_dt1
    Dim strcalcu_logic
    Dim strpay_logic
    Dim strEmp_no
    Dim strYear,strMonth,strDay

    Dim intCnt1
    Dim intCnt2
    Dim emp_no, strRetire_yymm

    Dim strMsg_cd
    Dim strMsg_text    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strretire_yyyy		= Request("txtretire_yyyy")
    strretire_strt_dt	= UNIConvDateCompanyToDB(Request("txtretire_strt_dt"),NULL)
    Call ExtractDateFrom(Request("txtretire_strt_dt"),gDateFormat,gComDateType,strYear,strMonth,strDay)
    strretire_strt_dt1	= strYear & strMonth & strDay
    strretire_end_dt	= UNIConvDateCompanyToDB(Request("txtretire_end_dt"),NULL)
    Call ExtractDateFrom(Request("txtretire_end_dt"),gDateFormat,gComDateType,strYear,strMonth,strDay)
    strretire_end_dt1	= strYear & strMonth & strDay
    strcalcu_logic		= Request("txtcalcu_logic")
    strpay_logic		= Request("txtpay_logic")


    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    End If


    if  strpay_logic = "D" then

        Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
        'lgStrSQL = " SELECT emp_no, retire_dt FROM haa010t"
        lgStrSQL = " SELECT emp_no, convert(varchar(8),retire_dt,112) retire_yymm FROM haa010t"
        lgStrSQL = lgStrSQL & " WHERE emp_no LIKE '" & stremp_no & "'"
        lgStrSQL = lgStrSQL & "   AND retire_dt >= '" & strretire_strt_dt & "'"
        lgStrSQL = lgStrSQL & "   AND retire_dt <= '" & strretire_end_dt & "'"

        IF  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true then
            Do While Not lgObjRs.EOF

                emp_no = ConvSPChars(lgObjRs("emp_no"))
                strRetire_yymm = Mid(lgObjRs("retire_yymm"), 1, 6)                
  
                If 	CommonQueryRs(" COUNT(*) "," HCA090T "," emp_no='" & emp_no & "' AND wk_yymm = '" & strRetire_yymm & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
                    intCnt1 = CInt(Replace(lgF0, Chr(11), ""))
                end if

                if  intCnt1 = 0 then
                    Call DisplayMsgBox("800377", vbInformation, "" , "", I_MKSCRIPT)
                    Call SetErrorStatus
                    Call SubCloseDB(lgObjConn)
                    Exit Sub
                end if
'급여가 있어도 퇴직금을 돌게 하기 위해 로직을 뺌 - 2003.9.4 by lsn                
'                If 	CommonQueryRs(" COUNT(*) "," HDF070T "," emp_no='" & emp_no & "' AND pay_yymm = '" & strRetire_yymm & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
 '                   intCnt2 = CInt(Replace(lgF0, Chr(11), ""))
  '              end if
   '             if  intCnt2 = 0 then
    '                Call DisplayMsgBox("800378", vbInformation, "", "", I_MKSCRIPT)
     '               Call SetErrorStatus
      '              Call SubCloseDB(lgObjConn)
       '             Exit Sub
        '        end if

    		    lgObjRs.MoveNext
               
            Loop 
        end if
        Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

    end if





    With lgObjComm
        .CommandText = "usp_hga050b1"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,  adXChar,adParamInput,Len(gUsrId), gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_cd",    adXChar,adParamInput,Len("R01"), "R01")
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cal_logic",    adXChar,adParamInput,Len(strcalcu_logic),   strcalcu_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_logic",    adXChar,adParamInput,Len(strPay_logic),     strPay_logic)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_strt_s",adXChar,adParamInput,Len(strretire_strt_dt1),strretire_strt_dt1)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_end_s", adXChar,adParamInput,Len(strretire_end_dt1), strretire_end_dt1)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",  adXChar,adParamInput,Len(strEmp_no),        strEmp_no)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,  adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,  adXChar,adParamOutput,60)

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

    Dim strretire_strt_dt,strretire_strt_dt1
    Dim strretire_end_dt,strretire_end_dt1
    Dim strEmp_no
    Dim strYear,strMonth,strDay

    Dim intCnt1
    Dim intCnt2
    Dim emp_no, strRetire_yymm

    Dim strMsg_cd
    Dim strMsg_text    
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strretire_strt_dt	= UNIConvDateCompanyToDB(Request("txtretire_strt_dt"),NULL)
    Call ExtractDateFrom(Request("txtretire_strt_dt"),gDateFormat,gComDateType,strYear,strMonth,strDay)
    strretire_strt_dt1	= strYear & strMonth & strDay


    strretire_end_dt	= UNIConvDateCompanyToDB(Request("txtretire_end_dt"),NULL)
    Call ExtractDateFrom(Request("txtretire_end_dt"),gDateFormat,gComDateType,strYear,strMonth,strDay)
    strretire_end_dt1	= strYear & strMonth & strDay


    strEmp_no = Request("txtEmp_no")
    If  strEmp_no = "" Then
        strEmp_no = "%"
    End If

'call svrmsgbox(lgOpModeCRUD &"/"& retire_cd &"/"&  retire_strt_s &"/"& retire_end_s & "/"& strEmp_no , vbinformation,i_mkscript) 

    With lgObjComm
        .CommandText = "usp_hga050b1_delete"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_cd",    adXChar,adParamInput,Len("R01"), "R01")
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_strt_s",adXChar,adParamInput,Len(strretire_strt_dt1),strretire_strt_dt1)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@retire_end_s", adXChar,adParamInput,Len(strretire_end_dt1), strretire_end_dt1)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",  adXChar,adParamInput,Len(strEmp_no), strEmp_no)
   	    
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,  adXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,  adXChar,adParamOutput,60)

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
