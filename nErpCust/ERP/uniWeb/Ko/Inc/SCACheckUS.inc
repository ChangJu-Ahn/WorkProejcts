<Script Language=VBScript>

     Dim PostString
     Dim objConn

     On Error Resume Next
     
     If parent.length > 0 Then
        If LCase(parent.item(0).name) <> "frmenu" Then
           top.location.href = "../scam.asp"
        End If
     Else
        top.location.href = "../scam.asp"
     End If    

     Set objConn = CreateObject("uniConnector.cGlobal")                
     PostString  = Trim(objConn.GetAspPostString)
     Set objConn = Nothing
     
     If PostString = "" Then
        top.location.href = "../scam.asp"
     End If    

</Script>

