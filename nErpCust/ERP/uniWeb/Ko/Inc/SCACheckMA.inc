<Script Language=VBScript>

     Dim PostString
     Dim objConn

     On Error Resume Next

     If LCase(parent.name) <> "frtoolbar" Then
        top.location.href = "../../scam.asp"
     End If


     Set objConn = CreateObject("uniConnector.cGlobal")                
     PostString  = Trim(objConn.GetAspPostString)
     Set objConn = Nothing
     
     If PostString = "" Then
        top.location.href = "../../scam.asp"
     End If    

</Script>

