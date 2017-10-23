<% 
Option Explicit
   Dim iADODBConnString
   Dim ClientIp
   Dim objFSO
   Dim MenuID
   Dim objFile
   Dim Timer
   Dim UsrID
    
    ClientIp = Request("ClientIp")
    Timer    = Request("Timer")
    MenuID   = Request("MenuID")
    UsrID    = Request("UsrID")
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile( "C:\log\" & Date & ".txt",8,True)
       
    objFile.WriteLine MenuID & ";" & Timer  & ";" & UsrID & ";" & ClientIp    
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

%>