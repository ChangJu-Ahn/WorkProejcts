<!-- #Include file="./CommResponse.inc" -->

<Script Language=VBScript Runat=Server>
'==============================================================================
'
'==============================================================================
Const VB_YES_NO_CANCEL = 35           'Button Count
Const VB_YES_NO        = 36           'Button Count
Const VB_OK_CANCEL     = 33           'Button Count

'==============================================================================
'
'==============================================================================
Const MSG_OK_STR       = "990000"      '☜:
Const MSG_DEADLOCK_STR = "999999"      '☜:
Const MSG_DBERROR_STR  = "999997"      '☜:

Const I_INSCRIPT = 0
Const I_MKSCRIPT = 1

Class FetchMsg
    Dim Severity
    Dim Text
    Dim CD
End Class

Function GetSvrDate(Byval pConnStr)

    dim za0008
	dim strDt
	
	On Error Resume Next

	Set za0008 = server.CreateObject("Za0008.Za0008GetSvrDt")
	
	za0008.ComCfg = pConnStr
	za0008.Execute

	strdt = za0008.ExportSvrDtServerDateSvrDt
	GetSvrDate = strdt   'Mid(strdt, 1, 10)
	
	Set za0008 = Nothing

End Function


Function LoginMsgBox(Byval iErrDesc,Byval iType,Byval iLoc)

    If iLoc = I_MKSCRIPT Then
		iErrDesc = Replace(iErrDesc, chr(34), chr(34) & chr(34))
		iErrDesc = Replace(iErrDesc, vbCrLf , chr(34) & " & vbCrLf & _ " & vbCrLf & vbTab & chr(34))
		
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "MsgBox " & chr(34) & iErrDesc & chr(34) & ", " & iType & ", " & chr(34) & Request.Cookies("unierp")("gLogoName") & " Login" & chr(34) & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
	
End Function


Function DisplayLoginMsg(ByVal pMsgId, ByVal pBtnKind, ByVal pMsg1, ByVal pMsg2, Byval iLoc, Byval plang, Byval pUsrId, Byval pConnStr)
	Dim tMsg
	Dim iCount

    Set tMsg = New FetchMsg
    
    tMsg.CD = ""
    tMsg.Severity = ""
    tMsg.Text = ""
    
    tMsg.CD = pMsgId

       
    If Len(Trim(pMsgId)) > 6 or Len(Trim(pMsgId)) < 1 Then
        If Len(pMsg1) > 0 Then
            tMsg.Text = pMsg1
        Else
            tMsg.Text = "Unknown error" & vbCrLf & "Error Code : " & pMsgId
        End If
        
        tMsg.Severity = "4"
    
    Else
        If MessageFetch(tMsg, plang, pConnStr) = True Then
        
            iCount = CountStrings(tMsg.Text, "%")
            If iCount > 0 Then
                tMsg.Text = MessageSplit(iCount, tMsg.Text, pMsg1, pMsg2)
            End If
        End If
        
    End IF
    
    If Instr(1, tMsg.Text, chr(34)) > 0 Then
		tMsg.Text = Replace(tMsg.Text, chr(34) , chr(34) & chr(34))
	End If
	If Instr(1, tMsg.Text, vbCrLf) > 0 Then
		tMsg.Text = Replace(tMsg.Text, vbCrLf , chr(34) & " & _ " & vbCrLf & chr(34))
	End If
    
    'If IsMissing(pBtnKind) Then
    If CStr(pBtnKind) = "0" or CStr(pBtnKind) = "64" Then
        Select Case tMsg.Severity
            Case "3"   ' Error
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbExclamation, iLoc)
            Case "4"   ' Fatal
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbCritical, iLoc)
            Case "2"   ' Warning
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbExclamation, iLoc)
            Case "1"   ' Information
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbInformation, iLoc)
        End Select

    Else
        Select Case pBtnKind
            Case 33   ' Ok, Cancel
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbOKCancel + vbQuestion, iLoc)
            Case 35   ' Yes,No,Cancel
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbYesNoCancel + vbQuestion, iLoc)
            Case 36   ' Yes,No
                DisplayLoginMsg = LoginMsgBox(tMsg.Text, vbYesNo + vbQuestion, iLoc)
            Case Else
				DisplayLoginMsg = LoginMsgBox("Check Button Option " & vbCrLf & "Button Code : " & pBtnKind, vbInformation, iLoc)
        End Select

    End If
    
'    If gSeverity <= tMsg.Severity Then
'        If MessageWrite(tMsg, pUsrId, pConnStr) = False Then
            'DisplayLoginMsg = LoginMsgBox("Message Logging Failed ... " & vbCrLf & tMsg.Text, vbInformation, iLoc)
'        End If
'    End If
    
	If CInt(gSeverity) <= CInt(tMsg.Severity) Then       'Level for Message logging
	   MessageWrite tMsg, pUsrId, pConnStr
	End If

End Function

Function MessageFetch(ByRef tMsg, Byval plang, Byval pConnStr)

    Dim pB1c039
    On Error Resume Next

    MessageFetch = True

	Set pB1c039 = Server.CreateObject("B1c039.B1c039LookupMessage")

    If Err.Number <> 0 Then
       tMsg.Text = Err.description
       tMsg.CD = Err.Number
       tMsg.Severity = "4"
       Err.Clear                                                        '☜: Clear error no
       Set pB1c039 = Nothing
       Exit Function
	End If
		
	'If Len(gLang) < 1 Then
	'	gLang = "KO"
	'End If

	pB1c039.ImportBMessageMsgCd = tMsg.CD
	pB1c039.ImportBLanguageLangCd = plang

    pB1c039.ComCfg = pConnStr
	pB1c039.Execute
    
    If Not (pB1c039.OperationStatusMessage = MSG_OK_STR) Then
		tMsg.Text = tMsg.CD & " : Message is not registered!"
		tMsg.Severity = "2"
		Set pB1c039 = Nothing
		Exit Function
    End If

    tMsg.Text = pB1c039.ExportBMessageMsgText
    tMsg.Severity = pB1c039.ExportBMessageSeverity

    Set pB1c039 = Nothing
	
End Function

Function MessageSplit(ByVal iCount, ByVal MsgText, _
                              ByVal pMsg1, ByVal pMsg2)
	Dim strMessage
	Dim strTemp
	Dim strMsgA
	Dim strMsgB

	Dim lColPos1
	Dim lColPos2
    
    strMessage = "" : strTemp = ""  
    strMsgA = "" : strMsgB = ""
    lColPos1 = 0 : lColPos2 = 0
        
    lColPos1 = InStr(1, MsgText, "%")

    strTemp = Mid(MsgText, lColPos1 + 1, 1)

    Select Case strTemp
        Case "1"
            strMessage = pMsg1
        Case "2"
            strMessage = pMsg2
    End Select

    If iCount = 2 Then

        lColPos2 = InStr(lColPos1 + 2, MsgText, "%")

        strMsgA = Mid(MsgText, lColPos1 + 2, lColPos2 - (lColPos1 + 2))

        strMessage = strMessage & " " & strMsgA

        strTemp = Mid(MsgText, lColPos2 + 1, 1)

        Select Case strTemp
            Case "1"
                strMessage = strMessage & " " & pMsg1
            Case "2"
                strMessage = strMessage & " " & pMsg2
        End Select

        strMsgB = Mid(MsgText, lColPos2 + 2, Len(MsgText) - lColPos2)

        strMessage = strMessage & " " & strMsgB

    Else

        strMsgA = Mid(MsgText, lColPos1 + 2, Len(MsgText) - lColPos1)
        strMessage = strMessage & " " & strMsgA

    End If

    MessageSplit = strMessage

End Function

'=============================================================================
' Parameter      : pProgId -> ".asp" Name
'                  pUsrId -> user ID
'                  pMsgCd -> Message Code
'                  pClientNm -> 
'                  pClientIp -> 
' Description    : This Script query db table , b_mesaage
' Return Value   : strMessage
'=============================================================================
Function MessageWrite(ByRef tMsg, ByVal pUsrId, ByVal pConnStr)
	Dim strTemp
	Dim strVal
	Dim lColPos
	Dim pZa0041
	
	On Error Resume Next
    
    MessageWrite = False
    
    strTemp = "" : strVal = "" : lColPos = 0

    Set pZa0041 = Server.CreateObject("Za0041.Za0041ControlMsgLogging")    

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		tMsg.Text = Err.Number & " : " & Err.description
		tMsg.Severity = "4"
		MessageWrite = False
		Set pZa0041 = Nothing												'☜: ComProxy Unload
		Exit Function
	End If
	
    '--------------------
    'Data manipulate area
    '--------------------
    pZa0041.ImportZMsgLoggingProgId = "uniloginprocess"
    pZa0041.ImportUsrZUsrMastRecUsrId = pUsrId
    pZa0041.ImportZMsgLoggingMsgCd = tMsg.CD
    pZa0041.ImportZMsgLoggingMsg = tMsg.Text
    pZa0041.ImportZMsgLoggingClientNm = Request.ServerVariables("REMOTE_ADDR")
    pZa0041.ImportZMsgLoggingClientIp = Request.ServerVariables("REMOTE_ADDR")
    pZa0041.ImportZMsgLoggingInsrtUserId = pUsrId
    pZa0041.ImportZMsgLoggingUpdtUserId = pUsrId

    pZa0041.ComCfg = pConnStr        
    pZa0041.execute

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    If Err.Number <> 0 Then
		tMsg.Text = Err.Number & " : " & Err.Description
		tMsg.Severity = "4"
		MessageWrite = False
		Set pZa0041 = Nothing
		Exit Function 
    End If

    '-----------------------------------------
    'Com action result check area(DB,internal)
    '-----------------------------------------
    If Not (pZa0041.OperationStatusMessage = MSG_OK_STR) Then
		tMsg.Text = pZa0041.OperationStatusMessage
		tMsg.Severity = "4"
		MessageWrite = False
		Set pZa0041 = Nothing
		Exit Function
    End If

    pZa0041.Clear
    Set pZa0041 = Nothing
    
    MessageWrite = True
    
End Function

'=============================================================================
' Function Name  : CountStrings
' Parameter      : strString -> Message text
'                  strTarget -> "%"
' Description    : This function is counting "%" value
' Return Value   : "%" Count
'=============================================================================

Function CountStrings(ByVal strString, ByVal strTarget)

    Dim lPosition
    Dim lCount
   
    lPosition = 1
    
    Do While InStr(lPosition, strString, strTarget)
    
        lPosition = InStr(lPosition, strString, strTarget) + 1
        lCount = lCount + 1
    
    Loop
    
    CountStrings = lCount
   
End Function

'========================================================================================
' Function Name : ConvSPChars
' Function Desc : 문자열안의 "를 ""로 바꾼다.
'========================================================================================
Function ConvSPChars(strVal)
	ConvSPChars = Replace(strVal, """", """""")
End Function 

'=============================================================================
' Function Name  : RetSeverity
' Parameter      : pSeverity -> Severity
' Description    : This function is Severity Change
' Return Value   : Severity
'=============================================================================
Function RetSeverity(Byval pSeverity)

    RetSeverity = 0
    
    If pSeverity >= 0 And pSeverity <= 10 Then RetSeverity = 1
    If pSeverity >= 11 And pSeverity <= 16 Then RetSeverity = 2
    If pSeverity >= 17 And pSeverity <= 19 Then RetSeverity = 3
    If pSeverity >= 20 And pSeverity <= 25 Then RetSeverity = 4

End Function
</script>
