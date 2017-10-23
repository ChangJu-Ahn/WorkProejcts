<% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<%
Dim iMode
Dim iReturnStr
Dim gEnvInf
Dim gLang

iMode = Request("iMode")
gLang = Request("LangCD")
gEnvInf = Request("EnvInf")

Select Case iMode
    Case "GET"
        iReturnStr = GetMsg()
    Case "WRITE"
        iReturnStr = Cstr(MsgWrite())
End Select

Response.Write iReturnStr

Function GetMsg()
    Dim iMsgCD
    Dim iEnvInf
    
    iMsgCD = Request("MsgCd")
    GetMsg = GetMessageData(iMsgCD)
End Function

Function MsgWrite()
    Dim iMCode, iMSeverity, iMText, iEnvInf
    
    iMCode = Request("MsgCd")
    iMSeverity = Request("MsgSeverity")
    iMText = Request("MsgText")
    
    MsgWrite = MessageWrite(iMCode, iMSeverity, iMText)
    
End Function

Function GetMessageData(ByVal iMsgCd)

   Dim iRET
   Dim iStrSQL
   Dim tgEnvInf
   Dim iErrCode
   Dim iErrDesc
   Dim iBool

   On Error Resume Next
   
   Err.Clear
   
   tgEnvInf = Split(gEnvInf, Chr(12))
   
   iStrSQL = "Select MSG_TYPE,SEVERITY,MSG_TEXT FROM B_MESSAGE "
   iStrSQL = iStrSQL & " WHERE MSG_CD  = '" & iMsgCd & "'"
   iStrSQL = iStrSQL & " AND   LANG_CD = '" & tgEnvInf(1) & "'"
   
   If CommonLookUpRs(tgEnvInf(0), iStrSQL, iRET, iErrCode, iErrDesc, True) = True Then
      GetMessageData = "Y" & Chr(12) & iRET(2) & Chr(12) & iRET(1)
   Else
      If iErrCode = 0 Then
         Select Case UCase(Trim(tgEnvInf(1)))
             Case "KO"
                GetMessageData = "X" & Chr(12) & "서버에 문제가 있거나 해당하는 메세지 코드가 없습니다 " & vbCrLf & "메시지   코드: " & iMsgCd & Chr(12) & "5"
             Case Else
                GetMessageData = "X" & Chr(12) & "Server problems or message code not exists. " & vbCrLf & "Message Code: " & iMsgCd & Chr(12) & "5"
         End Select
      Else
         GetMessageData = "X" & Chr(12) & iErrDesc & Chr(12) & "4"
      End If
   End If
   
   iRET = Split(GetMessageData, Chr(12))
   
   If IsNumeric(tgEnvInf(7)) And IsNumeric(iRET(2)) Then
      If CInt(tgEnvInf(7)) <= CInt(iRET(2)) Then       'Message logging 하기 위한 level
         iBool = MessageWrite(iMsgCd, iRET(2), iRET(1))
      End If
   End If
    
End Function

Function MessageWrite(ByVal MCode, ByVal MSeverity, ByVal MText)
    
   Dim tgEnvInf
   Dim iStrSQL
   Dim iVarOutPut
   Dim iErrCode
   Dim iErrDesc
   Dim iLng
   Dim Temp1
   
   Err.Clear
   
   MessageWrite = True
   
   tgEnvInf = Split(gEnvInf, Chr(12))
   
   If tgEnvInf(8) = "ORACLE" Then
      Temp1 = " sysdate "
   Else
      Temp1 = " getdate() "
   End If

   iStrSQL = " SELECT USR_ID FROM Z_USR_MAST_REC WHERE USR_ID = '" & tgEnvInf(3) & "'"
   If CommonLookUpRs(tgEnvInf(0), iStrSQL, iVarOutPut, iErrCode, iErrDesc, True) = False Then ' Error or user id not exists
      MessageWrite = False
   Else  '사용자 아이디 있을 때 
      
      iStrSQL = " SELECT MSG_CD,MSG_TYPE,SEVERITY,MSG_TEXT,LANG_CD FROM B_MESSAGE "  'can not find message
      iStrSQL = iStrSQL & " WHERE MSG_CD  = '" & MCode & "'"
      iStrSQL = iStrSQL & " AND   LANG_CD = '" & tgEnvInf(1) & "'"
      
      If CommonLookUpRs(tgEnvInf(0), iStrSQL, iVarOutPut, iErrCode, iErrDesc, True) Then
         
         iStrSQL = " INSERT INTO Z_MSG_LOGGING " & _
                "(OCCUR_DT,MSG_TYPE,MSG_CD,MSG,USR_ID,SEVERITY,PROG_ID,CLIENT_NM,CLIENT_IP,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) " & _
                "VALUES (" & Temp1 & ", '" & iVarOutPut(1) & "', '" & iVarOutPut(0) & "', '" & MText & "', '" & _
                tgEnvInf(3) & "', '" & iVarOutPut(2) & "', '" & tgEnvInf(2) & "', '" & tgEnvInf(4) & "', '" & _
                tgEnvInf(5) & "', '" & tgEnvInf(6) & "'," & Temp1 & ", '" & tgEnvInf(6) & "', " & Temp1 & ") "
                         
         If CommonTxRs(tgEnvInf(0), iStrSQL, iLng, iErrCode, iErrDesc) = False Then
         End If
      Else 'Can not find message or error
         If iErrCode = 0 Then
            iStrSQL = " INSERT INTO Z_MSG_LOGGING " & _
                   "(OCCUR_DT,MSG_TYPE,MSG_CD,MSG,USR_ID,SEVERITY,PROG_ID,CLIENT_NM,CLIENT_IP,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) " & _
                   "VALUES (" & Temp1 & ", 'S', '" & MCode & "', '" & MText & "', '" & _
                   tgEnvInf(3) & "', '" & MSeverity & "', '" & tgEnvInf(2) & "', '" & tgEnvInf(4) & "', '" & _
                   tgEnvInf(5) & "', '" & tgEnvInf(6) & "', " & Temp1 & ", '" & tgEnvInf(6) & "'," & Temp1 & ") "
                            
            If CommonTxRs(tgEnvInf(0), iStrSQL, iLng, iErrCode, iErrDesc) = False Then
            End If
         End If
      End If
   End If
    
End Function


Function CommonLookUpRs(ByVal gADODBConnString, ByVal pvStrSQL, prVarArray, prLngErrCode, prStrErrDesc, ByVal pvGetOrNot)
    
    Dim adoCn
    Dim adoRs
    Dim strSQL
    Dim iTemp
    Dim oTemp
    Dim iLoop
    
    On Error Resume Next
    Err.Clear
    
    
    CommonLookUpRs = True
    prLngErrCode = 0
    prStrErrDesc = ""
    
    Set adoCn = Server.CreateObject("ADODB.Connection")
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    
    adoCn.Open gADODBConnString
    
    Set adoRs = adoCn.Execute(pvStrSQL, , adCmdText)
       
    If Not (adoRs.EOF And adoRs.BOF) Then
       If pvGetOrNot = True Then
          iTemp = adoRs.GetRows()
          ReDim oTemp(UBound(iTemp, 1))
          For iLoop = 0 To UBound(iTemp)
              oTemp(iLoop) = iTemp(iLoop, 0)
          Next
          prVarArray = oTemp
       End If
    Else
       CommonLookUpRs = False
    End If
    
    Call CloseAdoObject(adoRs)
    Call CloseAdoObject(adoCn)
    
    If Err.number <> 0 Then
        CommonLookUpRs = False
        prLngErrCode = Err.Number
        prStrErrDesc = Err.Description
    End If    
End Function

Function CommonTxRs(ByVal gADODBConnString, ByVal pvStrSQL, prLngRecordsAffected, prLngErrCode, prStrErrDesc)

    Dim adoCn
    
    On Error Resume Next
    Err.Clear
    
    
    CommonTxRs = True
    prLngErrCode = 0
    prStrErrDesc = ""
    
    Set adoCn = Server.CreateObject("ADODB.Connection")
    
    adoCn.Open gADODBConnString
    
    adoCn.Execute pvStrSQL, prLngRecordsAffected, adCmdText + adExecuteNoRecords
        
    Call CloseAdoObject(adoCn)
    
    If Err.number <> 0 Then    
        CommonTxRs = False
        prLngErrCode = Err.Number
        prStrErrDesc = Err.Description
    End If        
End Function

Sub CloseAdoObject(pObject)

    If Not (pObject Is Nothing) Then
       If pObject.State = adStateOpen Then
          pObject.Close
       End If
       Set pObject = Nothing
    End If
End Sub

Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    pPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "KO" & "\Log\1.txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile( pPath,8,True)
       
    objFile.WriteLine pLogData
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

End Sub
%>