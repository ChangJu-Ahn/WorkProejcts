<% session.CodePage=949 %>

<Script Language=VBScript RUNAT=Server>

Dim gADOConnStr, gCursorLocation

gCursorLocation = adUseServer
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubOpenDB(pObjConn)
    On Error Resume Next
    Err.Clear

	Set pObjConn = Server.CreateObject("ADODB.Connection")

	pObjConn.ConnectionString  = gADODBConnString
	pObjConn.CommandTimeout = 300
	pObjConn.Open

    If CheckSYSTEMError(Err,True) = True Then
    End If

End Sub

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubCloseDB(pObjConn)
    On Error Resume Next
    Err.Clear

	If Not (pObjConn is Nothing) Then	' Nothing 체크 
		If pObjConn.State = 1 then		' 연결중이면 연결 해재 
			pObjConn.Close
		End If
		Set pObjConn = Nothing
	End If

    If CheckSYSTEMError(Err,True) = True Then
    End If

End Sub

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------

Function FncOpenRs(pCRUD,lgObjConn,pRs,pSource,pCursorType,pLockType)
    Dim  iCursorType
    Dim  iLockType

    On Error Resume Next
    Err.Clear

 
	FncOpenRs = False

    Set pRs = Server.CreateObject("ADODB.Recordset")
	
    Select Case UCase(pCRUD)
       Case "C"
               iCursorType = adOpenDynamic
               iLockType   = adLockPessimistic
       Case "R"
               iCursorType = adOpenForwardOnly
               iLockType   = adLockReadOnly
       Case "U"
               iCursorType = adOpenDynamic
               iLockType   = adLockPessimistic
       Case "D"
               iCursorType = adOpenDynamic
               iLockType   = adLockPessimistic
       Case "B"
               iCursorType = adOpenDynamic
               iLockType   = adLockBatchOptimistic
       Case "P"
               iCursorType = pCursorType
               iLockType   = pLockType
               pRs.CursorLocation = gCursorLocation	' -- 2005/03/02 최영태 추가 
    End Select

    pRs.Open pSource,lgObjConn,iCursorType,iLockType

    If CheckSYSTEMError(Err,True) = True Then
		
    End If

    If CheckSQLError(pRs.ActiveConnection,True) = False Then
       If Not(pRs Is Nothing) Then
          If Not( pRs.EOF And  pRs.BOF) Then
             FncOpenRs  = True
          End If
       End If
    End If
End Function

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubCloseRs(pRs)
	On Error Resume Next
    If IsNull(pRs) Then
       Exit Sub
    End If
    If Not(pRs Is Nothing) then
       pRs.Close										'☆: 레코드셋 닫기 
	   Set pRs = Nothing								'☆: 레코드셋 파기 
	End If
End Sub

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubCreateCommandObject(pObjComm)
    On Error Resume Next
    Err.Clear

	Set pObjComm = Server.CreateObject("ADODB.Command")

	pObjComm.ActiveConnection  = gADODBConnString

    If CheckSYSTEMError(Err,True) = True Then
    End If

End Sub

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubCloseCommandObject(pObjComm)
    On Error Resume Next
    Err.Clear
	Set pObjComm = Nothing
End Sub

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Function FncRsExists(pRs)
    FncRsExists = True
    If pRs.EOF And pRs.BOF Then
       FncRsExists = False
    End If
End Function


'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Sub SubSkipRs(pRs,pMax)
    Dim iDx

    For iDx = 1 To pMax
        pRs.MoveNext
    Next

End Sub

Sub SaveErrorLog(objError)

End Sub

Sub OnTransactionCommit()
    Call CommonOnTransactionCommit()
End Sub

Sub OnTransactionAbort()
    Call CommonOnTransactionAbort()
	Call SaveErrorLog(Err)
End Sub

Function CommonQueryRs(SelectList,FromList,WhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    Dim lgStrSQL
    Dim iObjRs
    Dim iObjConn
    Dim iDx

    On Error Resume Next
    Err.Clear

    CommonQueryRs = False
    
    lgF0 = ""
    lgF1 = ""
    lgF2 = ""
    lgF3 = ""
    lgF4 = ""
    lgF5 = ""
    lgF6 = ""

    lgStrSQL = "Select " & SelectList & " From   " & FromList

    If Trim(WhereList) > "" Then
       lgStrSQL = lgStrSQL  & " Where  " & WhereList
    End If

    Call SubOpenDB(iObjConn)
    If 	FncOpenRs("R",iObjConn,iObjRs,lgStrSQL,"X","X") = False Then
        lgF0 = "X"
    Else
        While Not iObjRs.EOF
           If iObjRs.Fields.Count > 0 Then
              lgF0 = lgF0 & iObjRs(0) & Chr(11)
              If iObjRs.Fields.Count > 1 Then
                 lgF1 = lgF1 & iObjRs(1) & Chr(11)
                 If iObjRs.Fields.Count > 2 Then
                    lgF2 = lgF2 & iObjRs(2) & Chr(11)
                    If iObjRs.Fields.Count > 3 Then
                       lgF3 = lgF3 & iObjRs(3) & Chr(11)
                       If iObjRs.Fields.Count > 4 Then
                          lgF4 = lgF4 & iObjRs(4) & Chr(11)
                          If iObjRs.Fields.Count > 5 Then
                             lgF5 = lgF5 & iObjRs(5) & Chr(11)
                             If iObjRs.Fields.Count > 5 Then
                                lgF6 = lgF6 & iObjRs(6) & Chr(11)
                             End If  ' 6
                          End If  ' 5
                       End If  ' 4
                    End If  ' 3
                 End If  ' 2
              End If  ' 1
           End If  ' 0
           iObjRs.MoveNext
        Wend
    End If
    
    CommonQueryRs = True
    
    Call SubCloseRs(iObjRs)
    Call SubCloseDB(iObjConn)

End Function

Function FuncCodeName(intSW, MajorCd, MinorCd)
    Dim iSelectList
    Dim iFromList
    Dim iWhereList
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6


    Select Case intSW
        Case 1                                                  ' B_MAJOR
              iSelectList = " MINOR_NM "
              iFromList   = " B_MINOR  "
              iWhereList  = " MAJOR_CD = '" & MajorCd & "' AND MINOR_CD = '" & MinorCd & "'"

        Case 2                                                  ' BCB020T   : 부서코드 
              iSelectList = " DEPT_NM "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & MinorCd & "')"
              Else
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
              End If
        Case 3                                                  ' B_COUNTRY : 국적 
              iSelectList = " COUNTRY_NM "
              iFromList   = " B_COUNTRY  "
              iWhereList  = " COUNTRY_CD = '" & MinorCd & "'"

        Case 4                                                  ' B_COMPANY : 회사코드 
              iSelectList = " CO_NM "
              iFromList   = " B_COMPANY  "
              iWhereList  = " CO_CD = '" & MinorCd & "'"
        Case 5                                                  ' 내부부서코드 
              iSelectList = " INTERNAL_CD "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & MinorCd & "')"
              Else
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
              End If

        Case 6                                                  ' B_COMPANY : 회사코드 
              iSelectList = " BANK_NM "
              iFromList   = " B_BANK  "
              iWhereList  = " BANK_CD = '" & MinorCd & "'"

	End Select

    Call CommonQueryRs(iSelectList,iFromList,iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If 	UCase(lgF0) = "X" Then
        FuncCodeName = MinorCd
    Else
        lgF0 = Split(lgF0,Chr(11))
        FuncCodeName = lgF0(0)
    End If

End Function

</Script>