<% 
   'Compability : uniConnector 456
   
   Call PMain
   
   
Sub PMain()

    Dim Obj
    Dim Obj2
    Dim ExecuteSql
    Dim Lang
    Dim iPath
    Dim ExecuteSvrDLL
    Dim strC
    Dim strF
    Dim strCompany
    Dim strClientID
    Dim strConnString
    Dim strUserIDKind
    Dim strUserNM
    Dim strAPPVersion
    Dim strOSVersion
    Dim strClientNM
    Dim strSQL

    On Error Resume Next
    
    iPath = Request.ServerVariables("PATH_INFO")
    iPath = Split(iPath, "/")
    Lang = UCase(iPath(UBound(iPath) - 2))
    
    If Err.Number <> 0 Then
       Response.Write "The Error of creation of Server Daemon Process"
       Response.End
    End If

    If Err.Number = 0 Then
       strC = Request("Cmd")
       strF = Request("Flag")
       strCompany = Request("Company")
       strClientID = Request("ClientID")
       strConnString = Request("ConnString")
       strUserIDKind = Trim(Request("UserIDKind"))
       strUserNM = Request("UN")                       '2005-06-17
       strAPPVersion = Request("AV")                   '2005-06-17
       strClientNM = Request("CN")                     '2005-06-17
       strOSVersion = Request("OV")                    '2005-06-17
       strSQL = Request("SQL")

       If strUserIDKind = "" Then
          strUserIDKind = "U"
        End If

       If strC = "R" Then
          If strF = "S" Then
             Set Obj = CreateObject("uniStub.CX01")
'            ExecuteSvrDLL = Obj.WriteLoginStatus(strCompany, strClientID, strUserIDKind)
             ExecuteSvrDLL = Obj.WriteLoginStatus2(strCompany, strClientID, strUserIDKind, strUserNM, strAPPVersion, strClientNM, strOSVersion)                      '2005-06-17
             Set Obj = Nothing
          End If
          If strF = "E" Then
             Set Obj = CreateObject("uniStub.CX01")
             ExecuteSvrDLL = Obj.DeleteLoginStatus(strCompany, strClientID, strUserIDKind)
             Set Obj = Nothing
          End If
          If strF = "P" Then
             Set Obj2 = CreateObject("uniStubMessenger.CXM01")
             ExecuteSvrDLL = Obj2.IGetNotifierData()
             Set Obj2 = Nothing
          End If
          Response.Write ExecuteSvrDLL
       Else
          Set Obj = CreateObject("uniStub.CX01")
          ExecuteSql = Obj.XMLHTTPConnectDB(strConnString, strF, strSQL, strC, Lang)
          Set Obj = Nothing
          Response.ContentType = "text/xml"
          Response.Write ExecuteSql
      End If
    End If
    
    If Err.Number <> 0 Then
       Response.Write Err.Description
    End If
    Call WriteToLog("[-" & strF & "][" & Timer & "][" & strClientID & "][" & strUserNM & "][" & strAPPVersion & "][" & strClientNM & "][" & strOSVersion & "][" & ExecuteSvrDLL & "]")
    
End Sub
    
Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    Exit sub
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile("C:\GetPuniSvrConn2.log", 8, True)
       
    objFile.WriteLine "[" & Date & "][" & Time & "]" & pLogData
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

End Sub
   
%>
