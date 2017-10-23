<% 
   'Compability : uniConnector 456
   
   Call PMain
   
   
Sub PMain()

    Dim iuniStub
    Dim iuniStubMessenger
    Dim iResponseStr
    Dim strF
    Dim strAPPVersion
    Dim strOSVersion
    Dim strClientNM
    Dim strSKEY
    Dim strXGTXV
    
    On Error Resume Next
    
    If Err.Number = 0 Then
       
       strF          = Request("Flag")
       strSKEY       = Request("SKEY")
       strAPPVersion = Request("AV")                   '2005-06-17
       strClientNM   = Request("CN")                     '2005-06-17
       strOSVersion  = Request("OV")                    '2005-06-17
       strXGTXV      = Request("XGTXV")                    '2005-06-17

       If strF = "S" Then
          Set iuniStub = CreateObject("uniStub.CX01")
          iResponseStr = iuniStub.WriteLoginStatus4(strSKEY, strAPPVersion, strClientNM, strOSVersion,strXGTXV)                     '2005-06-17
          Set iuniStub = Nothing
       End If
       
       If strF = "E" Then
          Set iuniStub = CreateObject("uniStub.CX01")
          iResponseStr = iuniStub.DeleteLoginStatus4(strSKEY)
          Set iuniStub = Nothing
       End If
       
       If strF = "P" Then
          Set iuniStubMessenger = CreateObject("uniStubMessenger.CXM01")
          iResponseStr = iuniStubMessenger.IGetNotifierData()
          Set iuniStubMessenger = Nothing
       End If
          
       Response.Write iResponseStr
          
    End If
    
    If Err.Number <> 0 Then
       Response.Write Err.Description
    End If
    
    Call WriteToLog("[" & strF & "][" & Timer & "][" & strSKEY & "][" & strXGTXV & "][" & strAPPVersion & "][" & strClientNM & "][" & strOSVersion & "][" & iResponseStr & "]")
    
End Sub
    
Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    Exit Sub
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile("C:\ZGetPuniSvrConn3.log", 8, True)
       
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
