<% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<%

   Dim docReceived, iTempNode
   Dim gFunctionCase, gLangCD, gCompany, gDBServer, gDatabase, gUsrID, gPgmID, gSpreadName
   Dim gFileName, gNodeXml, gMajorVersion, gSuperUser
   Dim gXMLPathForSuperUser

   On Error Resume Next

   Set docReceived = Server.CreateObject("MSXML2.DOMDocument.4.0")
   docReceived.async = False
   docReceived.load Request

   For Each iTempNode In docReceived.documentElement.childNodes
       Select Case iTempNode.TagName
           Case "Case"       : gFunctionCase = iTempNode.text
           Case "LangCD"     : gLangCD       = iTempNode.text
           Case "Company"    : gCompany      = iTempNode.text
           Case "DBServer"   : gDBServer     = iTempNode.text
           Case "Database"   : gDatabase     = iTempNode.text
           Case "UsrID"      : gUsrID        = iTempNode.text
           Case "PgmID"      : gPgmID        = iTempNode.text
           Case "SpreadName" : gSpreadName   = iTempNode.text
           Case "Data"       : gNodeXml      = iTempNode.FirstChild.xml
           Case "Version"    : gMajorVersion = iTempNode.text   
           Case "SuperUser"  : gSuperUser    = iTempNode.text                    
       End Select
   Next

   Set docReceived = Nothing

   gFileName            = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLangCD & "\XML\[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gUsrID & "]["     & gPgmID & "].xml"
   gXMLPathForSuperUser = Request.ServerVariables("APPL_PHYSICAL_PATH") & gLangCD & "\XML\[" & gCompany & "][" & gDBServer & "][" & gDatabase & "][" & gSuperUser & "][" & gPgmID & "].xml"

   Select Case gFunctionCase
      Case "Reset" : Call ResetNode()
      Case "Save"  : Call SaveNode()
      Case "Check" : Response.Write SvrXMLCheck()    
   End Select

Function SvrXMLCheck()

    Dim izRCXMLA00001C
    
    Set izRCXMLA00001C = CreateObject("zRCXML.A00001")
    SvrXMLCheck = izRCXMLA00001C.ISvrXMLCheck(gFileName,gSpreadName,gMajorVersion,gSuperUser,gUsrID,gXMLPathForSuperUser)
    Set izRCXMLA00001C = Nothing

End Function

Sub ResetNode()

    Dim izRCXMLA00001A
    
    Set izRCXMLA00001A = CreateObject("zRCXML.A00001")
    Call izRCXMLA00001A.IResetNode(gFileName,gSpreadName)
    Set izRCXMLA00001A = Nothing

End Sub

Sub SaveNode()

    Dim izRCXMLA00001
    
    Set izRCXMLA00001 = CreateObject("zRCXML.A00001")
    Call izRCXMLA00001.ISaveNode(gNodeXml,gFileName,gSpreadName)
    Set izRCXMLA00001 = Nothing
    
End Sub
%>