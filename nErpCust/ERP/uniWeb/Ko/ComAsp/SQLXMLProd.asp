ï»?% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<%

    Response.Expires=0
    Response.ContentType = "text/xml"
    
    Response.Write CommonQueryRs(Trim(Request("ADODBConnString")),Trim(Request("StrSQL")))


Function CommonQueryRs(ByVal gADODBConnString , ByVal pvStrSQL)
    Dim pRec
    
    On Error Resume Next
    
    CommonQueryRs = True

    Set pRec = Server.CreateObject("ADODB.Recordset")

    pRec.Open pvStrSQL, gADODBConnString
    
    CommonQueryRs = "<a>" & pvStrSQL & "</a>"

    If Not (pRec.EOF And pRec.BOF) Then
       CommonQueryRs = pRec(0)
    Else
       CommonQueryRs = False
    End If
    
    pRec.Close
    
    Set pRec = Nothing
    
End Function

%>



