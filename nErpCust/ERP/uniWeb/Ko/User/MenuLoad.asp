<!-- #Include file="../inc/CommResponse.inc" -->
<%
    Dim gConnect
    Dim StrUsr
    Dim UID
    Dim pRec
    Dim gLang
    
    Dim gPLProcessB
    Dim gPLProcessU
    
    StrUsr = Trim(Request.ServerVariables("HTTP_SOAPAction"))
    
    gConnect = Trim(Request.ServerVariables("HTTP_Connect"))
    
    If StrUsr = "GetDate" Then
        
        UID = Trim(Request.ServerVariables("HTTP_uid"))
        
        Set pRec = Server.CreateObject("ADODB.RecordSet")
            
        pRec.Open "SELECT CONVERT(CHAR(21),MAX(GEN_DT),20) FROM Z_AUTH_GEN WHERE USR_ID ='" & UID & "'", gConnect
            
        Response.Write pRec(0)

        pRec.Close
        
        Set pRec = Nothing
        
    ElseIf StrUsr = "BIZ" Then

        UID = Trim(Request.ServerVariables("HTTP_uid"))
        gLang = Trim(Request.ServerVariables("HTTP_lang"))
                
        Set gPLProcessB = Server.CreateObject("PLProcess.LCProcess")

        Response.Write gPLProcessB.PMakeBIZMenuXML(gConnect, UID, gLang) 'PLProcess.dll 18¿¡ PMakeMenuXML Ãß°¡µÊ.
        
        Set gPLProcessB = Nothing
    
    ElseIf StrUsr = "USR" Then
        
        UID = Trim(Request.ServerVariables("HTTP_uid"))
        gLang = Trim(Request.ServerVariables("HTTP_lang"))
        
        Set gPLProcessU = Server.CreateObject("PLProcess.LCProcess")

        Response.Write gPLProcessU.PMakeUSRMenuXML(gConnect, UID, gLang) 'PLProcess.dll 18¿¡ PMakeMenuXML Ãß°¡µÊ.
        
        Set gPLProcessU = Nothing            
        
    End If
    
	
%>


	