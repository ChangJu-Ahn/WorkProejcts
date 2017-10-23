<%
	Response.ContentType = "text/html"


Response.write GetGlobalXML(Request.Cookies("unierp")("gXMLFileNm"))%>

<%
Function GetGlobalXML(pXMLPath)   '2003-08-07 leejinsoo
on error resume next
    Dim EDCodeComEDCodeObj1
    Dim xmlDOMDocument

    Set xmlDOMDocument = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDOMDocument.async = False 
	    
	xmlDOMDocument.Load (pXMLPath)

    Set EDCodeComEDCodeObj1 = Server.CreateObject("uni2kCommon.ConnectorControl")
    GetGlobalXML = Replace(EDCodeComEDCodeObj1.DeCodeData(xmlDOMDocument.firstChild.firstChild.xml),vbCrLf,"")
    Set xmlDOMDocument      = Nothing
    Set EDCodeComEDCodeObj1 = Nothing
End Function
%>