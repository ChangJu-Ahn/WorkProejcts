<Script Language=VBScript>

     Dim PostString
     Dim objConn

     On Error Resume Next

     If typename( window.external) <> "HTMLDialog" Then
        document.location.href = "../../scam.asp"
     End If


     Set objConn = CreateObject("uniConnector.cGlobal")                
     PostString  = Trim(objConn.GetAspPostString)
     Set objConn = Nothing
     
     If PostString = "" Then
        document.location.href = "../../scam.asp"
     End If    

</Script>

