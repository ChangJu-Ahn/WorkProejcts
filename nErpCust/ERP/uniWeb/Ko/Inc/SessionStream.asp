<%

Response.Write GetSessionStream

Function GetSessionStream()

    Dim xmlDoc
    Dim xSessionDll
    
    On Error Resume Next

    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	Set xSessionDll = Server.CreateObject("xSession.A00001")
	xmlDoc.async = False 
	GetSessionStream = xSessionDll.DMakeDic(Request.Cookies("unierp")("SessionKey"),False)	
	Set xSessionDll = Nothing
    Set xmlDoc      = Nothing

End Function


%>