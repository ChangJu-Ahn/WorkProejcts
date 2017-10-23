<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    Dim intRetCD
    Dim strCalType
    Dim strCalDt
    Dim strWeekHoly
    Dim strSatHoly
    Dim strDayHoly
    Dim strUserId
	Dim strHolyRemark
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

     Call SubCreateCommandObject(lgObjComm)
	 Call SubMakeParameter()
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
'============================================================================================================
' Name : SubMakeParameter
' Desc : Make SP Parameter
'============================================================================================================
Sub SubMakeParameter()
	
	Dim LngRow
	Dim arrVal, arrTemp														'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus															'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim strSpace
	
	strCalType = UCase(Trim(Request("txtClnrType")))
	strCalDt = UNIConvYYYYMMDDToDate(gAPDateFormat,Trim(Request("txtYear")),"01","01")
    'strCalDt = Trim(Request("txtYear")) & "-01-01"
    strUserId = Trim(Request("txtUpdtUserId"))
    
    '-----------------------
    '요일 체크 
    '-----------------------
	
	If IsEmpty(Request("ChkSun")) = False Then
		strWeekHoly = "1"
	Else
		strWeekHoly = "0"
	End If
	If IsEmpty(Request("ChkMon")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else 
		strWeekHoly = strWeekHoly & "0"
	End If
	If IsEmpty(Request("ChkTue")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else 
		strWeekHoly = strWeekHoly & "0"
	End If
	If IsEmpty(Request("ChkWed")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else 
		strWeekHoly = strWeekHoly & "0"
	End If
	If IsEmpty(Request("ChkThu")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else  
		strWeekHoly = strWeekHoly & "0"
	End If
	If IsEmpty(Request("ChkFri")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else 
		strWeekHoly = strWeekHoly & "0"
	End If
	If IsEmpty(Request("ChkSat")) = False Then
		strWeekHoly = strWeekHoly & "1"
	Else 
		strWeekHoly = strWeekHoly & "0"
	End If

	'-----------------------------------
	'토요일 선택사항 체크(차후 추가사항)
	'-----------------------------------
	If IsEmpty(Request("ChkSat")) Then
		If IsEmpty(Request("ChkFirstWeek")) = False Then
			strSatHoly = "1"
		Else
			strSatHoly = "0"
		End If
		If IsEmpty(Request("ChkSecondWeek")) = False Then
			strSatHoly = strSatHoly & "1"
		Else
			strSatHoly = strSatHoly & "0"
		End If
		If IsEmpty(Request("ChkThirdWeek")) = False Then
			strSatHoly = strSatHoly & "1"
		Else
			strSatHoly = strSatHoly & "0"
		End If
		If IsEmpty(Request("ChkForthWeek")) = False Then
			strSatHoly = strSatHoly & "1"
		Else
			strSatHoly = strSatHoly & "0"
		End If
		If IsEmpty(Request("ChkFifthWeek")) = False Then
			strSatHoly = strSatHoly & "1"
		Else
			strSatHoly = strSatHoly & "0"
		End If
	Else
		strSatHoly = "00000"
	End If

    '-----------------------
    '휴일 체크 
    '-----------------------
    lgLngMaxRow = CInt(Request("txtMaxRows"))
	arrTemp = Split(Request("txtSpread"), gRowSep)						
	
	strDayHoly = ""
	strHolyRemark = ""
	
	If lgLngMaxRow > 0 Then
		For LngRow = 1 To lgLngMaxRow
			
			arrVal = Split(arrTemp(LngRow-1), gColSep)
			
			strStatus = arrVal(1)														'☜: Row 의 상태 
			
			Select Case strStatus
			    Case "C"																'⊙: Calendar 생성시에만 
					strDayHoly = strDayHoly & UniConvDate(Trim(arrVal(2))) & "!"
					strHolyRemark = strHolyRemark & Trim(arrVal(3)) & Space(20 - Len(Trim(arrVal(3)))) & "!"
			End Select
		Next
	Else
		strDayHoly = Space(10) & "!"
		strHolyRemark = Space(20) & "!"
	End If    
End Sub     
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strMsg_cd
    Dim strMsg_text

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    With lgObjComm
        .CommandText = "usp_gen_mfg_calendar"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cal_type",	advarXchar,adParamInput,2, strCalType)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@cal_dt_s",	advarXchar,adParamInput,10, strCalDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@week_holy",	advarXchar,adParamInput,7, strWeekHoly)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sat_holy",	advarXchar,adParamInput,6, strSatHoly)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@day_holy",	advarXchar,adParamInput,4000, strDayHoly)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@holy_remark",advarXchar,adParamInput,8000, strHolyRemark)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamInput,13, gUsrID)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
			End If
            IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End if
    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
               .DbExecOk
           End If
      End with
   End If   
       
</Script>	
