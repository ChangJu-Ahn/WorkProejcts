<% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<%

Dim docReceived, iTempNode
Dim gFunctionCase, gLangCD, gCompany, gDBServer, gDatabase, gUsrID, gPgmID, gSpreadName
Dim gFileName, gNodeXml, gMajorVersion, gSuperUser

'On error resume next

Set docReceived = Server.CreateObject("MSXML2.DOMDocument.4.0")
docReceived.async = False
docReceived.load Request

For Each iTempNode In docReceived.documentElement.childNodes
    Select Case iTempNode.TagName
        Case "Case"
            gFunctionCase = iTempNode.text
        Case "LangCD"
            gLangCD = iTempNode.text
        Case "Company"
            gCompany = iTempNode.text
        Case "DBServer"
            gDBServer = iTempNode.text
        Case "Database"
            gDatabase = iTempNode.text
        Case "UsrID"
            gUsrID = iTempNode.text
        Case "PgmID"
            gPgmID = iTempNode.text
        Case "SpreadName"
            gSpreadName = iTempNode.text
        Case "Data"
            gNodeXml = iTempNode.FirstChild.xml
        Case "Version"
            gMajorVersion = iTempNode.text   
        Case "SuperUser"
            gSuperUser = iTempNode.text                    
    End Select
Next

Set docReceived = Nothing

gFileName = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLangCD & "\XML\[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gUsrID & "][" & gPgmID & "].xml"

Select Case gFunctionCase
    Case "Reset"
        Call ResetNode()
    Case "Save"
        Call SaveNode()
    Case "Check"
        Response.Write SvrXMLCheck()    
End Select

Function SvrXMLCheck()
    Dim iDOM, iNode, iSvrVersion
   
    SvrXMLCheck = ""

    Set iDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
    iDOM.async = False
    Call iDOM.Load(gFileName)
    If iDOM.parseError.errorCode = 0 Then
        Set iNode = iDOM.selectSingleNode("/Root/" & gSpreadName)
        If TypeName(iNode) <> "Nothing" Then
            iSvrVersion = iNode.getAttribute("V")
            If iSvrVersion < gMajorVersion Then
                Call iDOM.documentElement.removeChild(iNode)
                Call iDOM.save(gFileName)
            Else
                SvrXMLCheck = iNode.xml
                Set iDOM = Nothing
                Exit Function
            End If
        End If
    End If
    If gSuperUser <> gUsrID Then
       SvrXMLCheck = LoadDefaultXML()
    End If
    
    Set iDOM = Nothing
    
End Function

Function LoadDefaultXML()
    Dim iDOM, iNode, iFileName, iSvrVersion
   
    LoadDefaultXML = ""
    
    iFileName = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLangCD & "\XML\[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gSuperUser & "][" & gPgmID & "].xml"
    Set iDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
    iDOM.async = False
    Call iDOM.Load(iFileName)
    If iDOM.parseError.errorCode = 0 Then
        Set iNode = iDOM.selectSingleNode("/Root/" & gSpreadName)
        If TypeName(iNode) <> "Nothing" Then
            iSvrVersion = iNode.getAttribute("V")
            If iSvrVersion = gMajorVersion Then
                LoadDefaultXML = iNode.xml
            End If
        End If
    End If
    
    Set iDOM = Nothing
    
End Function

Sub ResetNode()
    Dim iTempDOM, iTempNode

    Set iTempDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
    iTempDOM.async = False
    
    Call iTempDOM.Load(gFileName)
    If iTempDOM.parseError.errorCode = 0 Then
        Set iTempNode = iTempDOM.selectSingleNode("/Root/" & gSpreadName)
        If TypeName(iTempNode) <> "Nothing" Then
            Call iTempDOM.documentElement.removeChild(iTempNode)
            Call iTempDOM.save(gFileName)
        End If
    End If
    Set iTempDOM = Nothing
End Sub

Sub SaveNode()
    Dim iDOM, iTempDOM, iNode
    
    On Error Resume Next

    Set iDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
    Set iTempDOM = Server.CreateObject("MSXML2.DOMDocument.4.0")
   
    iTempDOM.async = False
    Call iTempDOM.loadXML(gNodeXml)
    If iTempDOM.parseError.errorCode = 0 Then
        iDOM.async = False
        Call iDOM.Load(gFileName)
        If iDOM.parseError.errorCode <> 0 Then
            Set iNode = iDOM.createElement("Root")
            Set iNode = iDOM.appendChild(iNode)
        Else
            Set iNode = iDOM.selectSingleNode("/Root/" & gSpreadName)
            If TypeName(iNode) <> "Nothing" Then
                Call iDOM.documentElement.removeChild(iNode)
            End If
        End If
        Call iDOM.documentElement.appendChild(iTempDOM.documentElement)
        Call iDOM.save(gFileName)
    End If

    Set iDOM     = Nothing
    Set iTempDOM = Nothing

End Sub
%>